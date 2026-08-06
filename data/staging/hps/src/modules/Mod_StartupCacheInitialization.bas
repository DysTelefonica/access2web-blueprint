Attribute VB_Name = "Mod_StartupCacheInitialization"
Option Compare Database
Option Explicit

Private Const CACHE_STEP_HPS_FULL As String = "HPS_FULL"
Private Const CACHE_STEP_HPS_INCREMENTAL As String = "HPS_INCREMENTAL"
Private Const CACHE_STEP_HPS_DERIVED_DRIFT As String = "HPS_DERIVED_DRIFT_TbHPS_HPS_SOLICITUDES_TbUsuarios"
Private Const CACHE_STEP_SICA_FULL As String = "SICA_FULL"
Private Const CACHE_STEP_SICA_CONTROLLED_FULL As String = "SICA_CONTROLLED_FULL_CAVEAT"
Private Const CACHE_STEP_INDICATORS_REFRESH As String = "INDICATORS_REFRESH"
Private Const CACHE_STEP_FAST_SYNC As String = "FAST_SYNC"
Private Const CACHE_STEP_COUNTERS_REFRESH As String = "COUNTERS_REFRESH"
Private Const CACHE_STEP_SUBFORM_SAFE_REQUERY As String = "SUBFORM_SAFE_REQUERY"
Private Const CACHE_STEP_STARTUP_FAST_SYNC As String = "STARTUP_FAST_SYNC"
Private Const CACHE_STEP_FORM_ACCESS_ENABLE As String = "FORM_ACCESS_ENABLE"
Private Const CACHE_STEP_HISTORICAL_FULL_REBUILD As String = "HISTORICAL_FULL_REBUILD"

Public Function EjecutarSincronizacionCargaInicialFrontend( _
    Optional ByRef p_Error As String _
) As String

    On Error GoTo errores

    p_Error = vbNullString
    EjecutarSincronizacionCargaInicialFrontend = SincronizarCachesLocalesFrontend(p_Error:=p_Error)
    If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
    Exit Function

errores:
    If p_Error = vbNullString Then
        If Err.Number = 1000 Then
            p_Error = Err.Description
        Else
            p_Error = "EjecutarSincronizacionCargaInicialFrontend: " & Err.Number & " - " & Err.Description
        End If
    End If
    EjecutarSincronizacionCargaInicialFrontend = vbNullString
End Function

Public Function RegenerarUsuariosHistoricosLocalesFrontend( _
    Optional ByRef p_Error As String _
) As String

    On Error GoTo errores

    p_Error = vbNullString
    RegenerarUsuariosHistoricosLocalesFrontend = cache_usuarios_historicos_regenerar(p_Error:=p_Error)
    If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
    Exit Function

errores:
    If p_Error = vbNullString Then
        If Err.Number = 1000 Then
            p_Error = Err.Description
        Else
            p_Error = "RegenerarUsuariosHistoricosLocalesFrontend: " & Err.Number & " - " & Err.Description
        End If
    End If
    RegenerarUsuariosHistoricosLocalesFrontend = vbNullString
End Function

Public Function SincronizarCachesLocalesFrontend( _
    Optional ByVal p_ForzarRegeneracionTotal As EnumSiNo = EnumSiNo.No, _
    Optional ByRef p_Error As String _
) As String

    On Error GoTo errores

    p_Error = vbNullString

    If p_ForzarRegeneracionTotal <> EnumSiNo.No Then
        AvanceNuevo "Regenerando caché HPS local completa..."
        InicializarTablasLocalesYEntidad_Optimizado _
            p_ForzarRegeneracionTotal:=EnumSiNo.Sí, _
            p_Error:=p_Error
        If p_Error <> vbNullString Then Err.Raise 1000, , p_Error

        AvanceNuevo "Regenerando caché SICA local completa..."
        ActualizaUsuariosSICALocal p_Error:=p_Error
        If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
    Else
        AvanceNuevo "Sincronizando diferencias base de usuarios..."
        InicializarTablasLocalesYEntidad_Optimizado _
            p_ForzarRegeneracionTotal:=EnumSiNo.No, _
            p_Error:=p_Error
        If p_Error <> vbNullString Then Err.Raise 1000, , p_Error

        AvanceNuevo "Verificando deriva derivada HPS..."
        RefrescarDriftDerivadoHps p_Error:=p_Error
        If p_Error <> vbNullString Then Err.Raise 1000, , p_Error

        ' Caveat: SICA still has no cheap source-vs-local drift seam equivalent to HPS.
        ' Keep the step explicit and controlled instead of hiding it behind the startup path.
        AvanceNuevo "Sincronizando caché SICA local controlada..."
        ActualizaUsuariosSICALocal p_Error:=p_Error
        If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
    End If

    ' [PR1b] Refresh the historical-user cache at startup. Previously this
    ' was only triggered manually via RegenerarUsuariosHistoricosLocalesFrontend,
    ' which meant that on a fresh start the form_FormInicial08UsuariosHistoricosDatos
    ' search form showed a stale cache (historical users from prior moves that
    ' hadn't been synced to TbUsuariosHistoricosLocal yet). A full regen
    ' at startup is cheap (the historical table is bounded by the number of
    ' actual moves, not the total user count) and gives operators a clean
    ' search view from second 0. Best-effort: if it fails, the operator
    ' can still trigger a manual regen from the form.
    AvanceNuevo "Regenerando caché de usuarios históricos completa..."
    cache_usuarios_historicos_regenerar p_Error:=p_Error
    If p_Error <> vbNullString Then Err.Raise 1000, , p_Error

    RefrescarCopiasIndicadores p_Error:=p_Error
    If p_Error <> vbNullString Then Err.Raise 1000, , p_Error

    SincronizarCachesLocalesFrontend = "OK"
    Exit Function

errores:
    If p_Error = vbNullString Then
        If Err.Number = 1000 Then
            p_Error = Err.Description
        Else
            p_Error = "SincronizarCachesLocalesFrontend: " & Err.Number & " - " & Err.Description
        End If
    End If
    SincronizarCachesLocalesFrontend = vbNullString
End Function

Public Function RegenerarCachesLocalesFrontend( _
    Optional ByVal p_ForzarRegeneracionTotal As EnumSiNo = EnumSiNo.Sí, _
    Optional ByRef p_Error As String _
) As String

    On Error GoTo errores

    p_Error = vbNullString
    RegenerarCachesLocalesFrontend = SincronizarCachesLocalesFrontend( _
        p_ForzarRegeneracionTotal:=p_ForzarRegeneracionTotal, _
        p_Error:=p_Error)
    If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
    Exit Function

errores:
    If p_Error = vbNullString Then
        If Err.Number = 1000 Then
            p_Error = Err.Description
        Else
            p_Error = "RegenerarCachesLocalesFrontend: " & Err.Number & " - " & Err.Description
        End If
    End If
    RegenerarCachesLocalesFrontend = vbNullString
End Function

Public Function DescribirPlanSincronizarCachesLocalesFrontend( _
    Optional ByVal p_ForzarRegeneracionTotal As EnumSiNo = EnumSiNo.No, _
    Optional ByRef p_Error As String _
) As String

    On Error GoTo errores

    p_Error = vbNullString

    If p_ForzarRegeneracionTotal = EnumSiNo.No Then
        DescribirPlanSincronizarCachesLocalesFrontend = _
            CACHE_STEP_HPS_INCREMENTAL & "|" & _
            CACHE_STEP_HPS_DERIVED_DRIFT & "|" & _
            CACHE_STEP_SICA_CONTROLLED_FULL & "|" & _
            CACHE_STEP_INDICATORS_REFRESH
    Else
        DescribirPlanSincronizarCachesLocalesFrontend = _
            CACHE_STEP_HPS_FULL & "|" & _
            CACHE_STEP_SICA_FULL & "|" & _
            CACHE_STEP_INDICATORS_REFRESH
    End If
    Exit Function

errores:
    p_Error = "DescribirPlanSincronizarCachesLocalesFrontend: " & Err.Number & " - " & Err.Description
    DescribirPlanSincronizarCachesLocalesFrontend = vbNullString
End Function

Public Function DescribirPlanCargaInicialFormInicial( _
    Optional ByRef p_Error As String _
) As String

    Dim syncPlan As String

    On Error GoTo errores

    p_Error = vbNullString

    syncPlan = DescribirPlanSincronizarCachesLocalesFrontend(p_Error:=p_Error)
    If p_Error <> vbNullString Then Err.Raise 1000, , p_Error

    DescribirPlanCargaInicialFormInicial = _
        CACHE_STEP_STARTUP_FAST_SYNC & "|" & _
        syncPlan & "|" & _
        CACHE_STEP_FORM_ACCESS_ENABLE
    Exit Function

errores:
    p_Error = "DescribirPlanCargaInicialFormInicial: " & Err.Number & " - " & Err.Description
    DescribirPlanCargaInicialFormInicial = vbNullString
End Function

Public Function DescribirPlanRegenerarUsuariosHistoricosLocalesFrontend( _
    Optional ByRef p_Error As String _
) As String

    On Error GoTo errores

    p_Error = vbNullString
    DescribirPlanRegenerarUsuariosHistoricosLocalesFrontend = CACHE_STEP_HISTORICAL_FULL_REBUILD
    Exit Function

errores:
    p_Error = "DescribirPlanRegenerarUsuariosHistoricosLocalesFrontend: " & Err.Number & " - " & Err.Description
    DescribirPlanRegenerarUsuariosHistoricosLocalesFrontend = vbNullString
End Function

Public Function DescribirPlanRegenerarCachesLocalesFrontend( _
    Optional ByVal p_ForzarRegeneracionTotal As EnumSiNo = EnumSiNo.Sí, _
    Optional ByRef p_Error As String _
) As String

    On Error GoTo errores

    p_Error = vbNullString
    DescribirPlanRegenerarCachesLocalesFrontend = DescribirPlanSincronizarCachesLocalesFrontend( _
        p_ForzarRegeneracionTotal:=p_ForzarRegeneracionTotal, _
        p_Error:=p_Error)
    If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
    Exit Function

errores:
    p_Error = "DescribirPlanRegenerarCachesLocalesFrontend: " & Err.Number & " - " & Err.Description
    DescribirPlanRegenerarCachesLocalesFrontend = vbNullString
End Function

Public Function DescribirPlanBotonPrincipalActualizarCachesFrontend( _
    Optional ByRef p_Error As String _
) As String

    On Error GoTo errores

    p_Error = vbNullString
    DescribirPlanBotonPrincipalActualizarCachesFrontend = _
        CACHE_STEP_FAST_SYNC & "|" & _
        CACHE_STEP_COUNTERS_REFRESH
    Exit Function

errores:
    p_Error = "DescribirPlanBotonPrincipalActualizarCachesFrontend: " & Err.Number & " - " & Err.Description
    DescribirPlanBotonPrincipalActualizarCachesFrontend = vbNullString
End Function

Public Function DescribirPlanBotonIndicadoresActualizarCachesFrontend( _
    Optional ByRef p_Error As String _
) As String

    On Error GoTo errores

    p_Error = vbNullString
    DescribirPlanBotonIndicadoresActualizarCachesFrontend = _
        CACHE_STEP_FAST_SYNC & "|" & _
        CACHE_STEP_INDICATORS_REFRESH & "|" & _
        CACHE_STEP_SUBFORM_SAFE_REQUERY
    Exit Function

errores:
    p_Error = "DescribirPlanBotonIndicadoresActualizarCachesFrontend: " & Err.Number & " - " & Err.Description
    DescribirPlanBotonIndicadoresActualizarCachesFrontend = vbNullString
End Function

Private Function RefrescarDriftDerivadoHps(Optional ByRef p_Error As String) As String
    Dim dbSource As DAO.Database
    Dim dbLocal As DAO.Database
    Dim ids As Object
    Dim idsHpsSource As Object
    Dim rsSource As DAO.Recordset
    Dim rsLocal As DAO.Recordset
    Dim idUsuario As String
    Dim minSource As Variant
    Dim minLocal As Variant
    Dim key As Variant

    On Error GoTo errores

    p_Error = vbNullString
    Set dbSource = getdb(p_Error)
    If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
    Set dbLocal = CurrentDb()
    Set ids = CreateObject("Scripting.Dictionary")
    ids.CompareMode = vbTextCompare
    Set idsHpsSource = CreateObject("Scripting.Dictionary")
    idsHpsSource.CompareMode = vbTextCompare

    Set rsSource = dbSource.OpenRecordset( _
        "SELECT IDUsuario, MIN(F_Concesion) AS MinConcesion " & _
        "FROM TbHPS WHERE IDUsuario Is Not Null GROUP BY IDUsuario", _
        dbOpenSnapshot)

    Do While Not rsSource.EOF
        idUsuario = Trim$(CStr(Nz(rsSource!idUsuario, vbNullString)))
        If idUsuario <> vbNullString Then
            AgregarIdDrift idsHpsSource, idUsuario
            minSource = rsSource!MinConcesion
            minLocal = FechaHpsMinimaLocal(dbLocal, idUsuario)
            If IsEmpty(minLocal) Then
                AgregarIdDrift ids, idUsuario
            ElseIf Not MismasFechasNulas(minSource, minLocal) Then
                AgregarIdDrift ids, idUsuario
            End If
        End If
        rsSource.MoveNext
    Loop

    rsSource.Close
    Set rsSource = Nothing

    Set rsLocal = dbLocal.OpenRecordset( _
        "SELECT ID FROM TbDatosLocal WHERE " & _
        "Not HPS_NAC_F_Concesion Is Null OR Not HPS_OTAN_F_Concesion Is Null OR " & _
        "Not HPS_ESA_F_Concesion Is Null OR Not HPS_UE_F_Concesion Is Null OR " & _
        "HPS_NAC_Activo='Sí' OR HPS_OTAN_Activo='Sí' OR HPS_ESA_Activo='Sí' OR HPS_UE_Activo='Sí'", _
        dbOpenSnapshot)

    Do While Not rsLocal.EOF
        idUsuario = Trim$(CStr(Nz(rsLocal!ID, vbNullString)))
        If idUsuario <> vbNullString Then
            If Not idsHpsSource.Exists(idUsuario) Then
                AgregarIdDrift ids, idUsuario
            End If
        End If
        rsLocal.MoveNext
    Loop

    rsLocal.Close
    Set rsLocal = Nothing

    For Each key In ids.Keys
        AvanceNuevo "Actualizando deriva HPS usuario ID: " & CStr(key)
        RefreshHpsUserCacheAfterMutation CStr(key), dbSource, dbLocal, p_Error
        If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
    Next key

    RefrescarDriftDerivadoHps = "OK"

SALIR:
    On Error Resume Next
    If Not rsSource Is Nothing Then rsSource.Close
    If Not rsLocal Is Nothing Then rsLocal.Close
    If Not dbSource Is Nothing Then dbSource.Close
    Set rsSource = Nothing
    Set rsLocal = Nothing
    Set dbSource = Nothing
    Set dbLocal = Nothing
    Set ids = Nothing
    Set idsHpsSource = Nothing
    Exit Function

errores:
    If p_Error = vbNullString Then
        If Err.Number = 1000 Then
            p_Error = Err.Description
        Else
            p_Error = "RefrescarDriftDerivadoHps: " & Err.Number & " - " & Err.Description
        End If
    End If
    RefrescarDriftDerivadoHps = vbNullString
    Resume SALIR
End Function

Private Sub RefrescarCopiasIndicadores(Optional ByRef p_Error As String)
    AvanceNuevo "Actualizando cachés de indicadores..."
    CopiarDatosAIndicadores p_Error:=p_Error, p_CacheAcabDeRegenerar:=True
    If p_Error <> vbNullString Then Err.Raise 1000, , p_Error

    CopiarDatosAIndicadores p_Error:=p_Error
    If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
End Sub

Private Function FechaHpsMinimaLocal(ByRef p_Db As DAO.Database, ByVal p_IDUsuario As String) As Variant
    Dim rs As DAO.Recordset

    Set rs = p_Db.OpenRecordset( _
        "SELECT FechaHPSConcesionMinima FROM TbDatosLocal WHERE ID=" & CLng(p_IDUsuario), _
        dbOpenSnapshot)

    If rs.EOF Then
        FechaHpsMinimaLocal = Empty
    Else
        FechaHpsMinimaLocal = rs!FechaHPSConcesionMinima
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Function MismasFechasNulas(ByVal p_A As Variant, ByVal p_B As Variant) As Boolean
    If IsNull(p_A) And IsNull(p_B) Then
        MismasFechasNulas = True
    ElseIf IsDate(p_A) And IsDate(p_B) Then
        MismasFechasNulas = (DateValue(CDate(p_A)) = DateValue(CDate(p_B)))
    Else
        MismasFechasNulas = False
    End If
End Function

Private Sub AgregarIdDrift(ByRef p_Ids As Object, ByVal p_IDUsuario As String)
    If Len(Trim$(p_IDUsuario)) = 0 Then Exit Sub
    If Not p_Ids.Exists(p_IDUsuario) Then p_Ids.Add p_IDUsuario, True
End Sub

