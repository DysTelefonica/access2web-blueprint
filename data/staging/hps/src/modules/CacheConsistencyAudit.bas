Attribute VB_Name = "CacheConsistencyAudit"
Option Compare Database
Option Explicit

Public Function RefreshHpsUserCacheAfterMutation( _
                                                ByVal p_IDUsuario As String, _
                                                Optional ByRef p_SourceDb As DAO.Database = Nothing, _
                                                Optional ByRef p_LocalDb As DAO.Database = Nothing, _
                                                Optional ByRef p_Error As String _
                                                ) As String

    Dim dbSource As DAO.Database
    Dim dbLocal As DAO.Database
    Dim ownsSource As Boolean
    Dim hpsCount As Long
    Dim minDate As Variant

    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_IDUsuario)) = 0 Then
        p_Error = "RefreshHpsUserCacheAfterMutation: missing IDUsuario"
        Err.Raise 1000
    End If

    If p_SourceDb Is Nothing Then
        Set dbSource = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
        ownsSource = True
    Else
        Set dbSource = p_SourceDb
    End If

    If p_LocalDb Is Nothing Then
        Set dbLocal = CurrentDb()
    Else
        Set dbLocal = p_LocalDb
    End If

    minDate = HpsMinConcessionDate(dbSource, p_IDUsuario)
    UpdateUserMinDate dbSource, p_IDUsuario, minDate

    hpsCount = CountRowsInDb(dbSource, "TbHPS", "IDUsuario=" & CLng(p_IDUsuario))
    If hpsCount = 0 Then
        DeleteHpsUserCacheRows dbSource, dbLocal, p_IDUsuario
    Else
        UpsertHpsUserCacheRow dbSource, dbSource, "TbUsuariosEntidades", p_IDUsuario, minDate
        UpsertHpsUserCacheRow dbSource, dbLocal, "TbDatosLocal", p_IDUsuario, minDate
        UpsertHpsUserCacheRow dbSource, dbLocal, "TbDatosLocalParaIndicadores", p_IDUsuario, minDate
    End If

    RefreshHpsUserCacheAfterMutation = "OK"

SALIR:
    If ownsSource Then
        If Not dbSource Is Nothing Then dbSource.Close
    End If
    Set dbSource = Nothing
    Set dbLocal = Nothing
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RefreshHpsUserCacheAfterMutation: " & Err.Description
    End If
    Resume SALIR
End Function

Public Function DeleteSicaLocalCaches( _
                                      ByVal p_IDSICA As String, _
                                      Optional ByRef p_SourceDb As DAO.Database = Nothing, _
                                      Optional ByRef p_LocalDb As DAO.Database = Nothing, _
                                      Optional ByRef p_Error As String _
                                      ) As String

    Dim dbLocal As DAO.Database

    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_IDSICA)) = 0 Then
        p_Error = "DeleteSicaLocalCaches: missing IDSICA"
        Err.Raise 1000
    End If

    If p_LocalDb Is Nothing Then
        Set dbLocal = CurrentDb()
    Else
        Set dbLocal = p_LocalDb
    End If

    DeleteSicaCacheRows dbLocal, p_IDSICA
    DeleteSicaLocalCaches = "OK"

SALIR:
    Set dbLocal = Nothing
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "DeleteSicaLocalCaches: " & Err.Description
    End If
    Resume SALIR
End Function

Public Function RefreshSicaLocalCaches( _
                                       ByVal p_IDSICA As String, _
                                       Optional ByRef p_SourceDb As DAO.Database = Nothing, _
                                       Optional ByRef p_LocalDb As DAO.Database = Nothing, _
                                       Optional ByRef p_Error As String _
                                       ) As String

    Dim dbSource As DAO.Database
    Dim dbLocal As DAO.Database
    Dim ownsSource As Boolean

    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_IDSICA)) = 0 Then
        p_Error = "RefreshSicaLocalCaches: missing IDSICA"
        Err.Raise 1000
    End If

    If p_SourceDb Is Nothing Then
        Set dbSource = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
        ownsSource = True
    Else
        Set dbSource = p_SourceDb
    End If

    If p_LocalDb Is Nothing Then
        Set dbLocal = CurrentDb()
    Else
        Set dbLocal = p_LocalDb
    End If

    DeleteSicaCacheRows dbLocal, p_IDSICA
    If CountRowsInDb(dbSource, "TbUsuariosSICA", "ID='" & SqlText(p_IDSICA) & "'") > 0 Then
        UpsertSicaCacheRow dbSource, dbLocal, "TbUsuariosSICALocal", p_IDSICA
        UpsertSicaCacheRow dbSource, dbLocal, "TbUsuariosSICALocalParaIndicadores", p_IDSICA
    End If

    RefreshSicaLocalCaches = "OK"

SALIR:
    If ownsSource Then
        If Not dbSource Is Nothing Then dbSource.Close
    End If
    Set dbSource = Nothing
    Set dbLocal = Nothing
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RefreshSicaLocalCaches: " & Err.Description
    End If
    Resume SALIR
End Function

Private Function HpsMinConcessionDate(ByRef p_Db As DAO.Database, ByVal p_IDUsuario As String) As Variant
    Dim rs As DAO.Recordset

    Set rs = p_Db.OpenRecordset( _
        "SELECT MIN(F_Concesion) AS MinConcesion FROM TbHPS WHERE IDUsuario=" & CLng(p_IDUsuario) & " AND F_Concesion Is Not Null", _
        dbOpenSnapshot)
    If Not rs.EOF Then
        HpsMinConcessionDate = rs!MinConcesion
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub UpdateUserMinDate(ByRef p_Db As DAO.Database, ByVal p_IDUsuario As String, ByVal p_MinDate As Variant)
    Dim sql As String

    If IsDate(p_MinDate) Then
        sql = "UPDATE TbUsuarios SET FechaHPSConcesionMinima=#" & Format$(CDate(p_MinDate), "mm/dd/yyyy") & "# WHERE ID=" & CLng(p_IDUsuario)
    Else
        sql = "UPDATE TbUsuarios SET FechaHPSConcesionMinima=Null WHERE ID=" & CLng(p_IDUsuario)
    End If
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub DeleteHpsUserCacheRows(ByRef p_SourceDb As DAO.Database, ByRef p_LocalDb As DAO.Database, ByVal p_IDUsuario As String)
    p_LocalDb.Execute "DELETE * FROM TbDatosLocal WHERE ID=" & CLng(p_IDUsuario), dbFailOnError
    p_LocalDb.Execute "DELETE * FROM TbDatosLocalParaIndicadores WHERE ID=" & CLng(p_IDUsuario), dbFailOnError
    p_SourceDb.Execute "DELETE * FROM TbUsuariosEntidades WHERE ID=" & CLng(p_IDUsuario), dbFailOnError
End Sub

Private Sub UpsertHpsUserCacheRow( _
                                  ByRef p_SourceDb As DAO.Database, _
                                  ByRef p_TargetDb As DAO.Database, _
                                  ByVal p_TableName As String, _
                                  ByVal p_IDUsuario As String, _
                                  ByVal p_MinDate As Variant)

    Dim rsTarget As DAO.Recordset
    Dim rsUser As DAO.Recordset
    Dim rsHps As DAO.Recordset
    Dim prefix As String

    Set rsUser = p_SourceDb.OpenRecordset("SELECT * FROM TbUsuarios WHERE ID=" & CLng(p_IDUsuario), dbOpenSnapshot)
    If rsUser.EOF Then
        DeleteHpsUserCacheRows p_SourceDb, p_TargetDb, p_IDUsuario
        GoTo SALIR
    End If

    Set rsTarget = p_TargetDb.OpenRecordset("SELECT * FROM " & p_TableName & " WHERE ID=" & CLng(p_IDUsuario), dbOpenDynaset)
    If rsTarget.EOF Then
        rsTarget.AddNew
        SetFieldValue rsTarget, "ID", CLng(p_IDUsuario)
    Else
        rsTarget.Edit
    End If

    CopyUserFields rsUser, rsTarget
    SetFieldValue rsTarget, "FechaHPSConcesionMinima", p_MinDate
    ClearHpsFields rsTarget

    Set rsHps = p_SourceDb.OpenRecordset("SELECT * FROM TbHPS WHERE IDUsuario=" & CLng(p_IDUsuario), dbOpenSnapshot)
    Do While Not rsHps.EOF
        prefix = HpsPrefix(CStr(Nz(rsHps!tipoHps, "")))
        If prefix <> "" Then CopyHpsFields rsHps, rsTarget, prefix
        rsHps.MoveNext
    Loop

    rsTarget.Update

SALIR:
    On Error Resume Next
    If Not rsHps Is Nothing Then rsHps.Close
    If Not rsTarget Is Nothing Then rsTarget.Close
    If Not rsUser Is Nothing Then rsUser.Close
    Set rsHps = Nothing
    Set rsTarget = Nothing
    Set rsUser = Nothing
End Sub

Private Sub DeleteSicaCacheRows(ByRef p_LocalDb As DAO.Database, ByVal p_IDSICA As String)
    p_LocalDb.Execute "DELETE * FROM TbUsuariosSICALocal WHERE ID='" & SqlText(p_IDSICA) & "'", dbFailOnError
    p_LocalDb.Execute "DELETE * FROM TbUsuariosSICALocalParaIndicadores WHERE ID='" & SqlText(p_IDSICA) & "'", dbFailOnError
End Sub

Private Sub UpsertSicaCacheRow(ByRef p_SourceDb As DAO.Database, ByRef p_LocalDb As DAO.Database, ByVal p_TableName As String, ByVal p_IDSICA As String)
    Dim rsSource As DAO.Recordset
    Dim rsTarget As DAO.Recordset
    Dim fld As DAO.Field

    ' Keep cache correctness independent from the external Expedientes supplier
    ' backend. The SICA cache consistency contract is about the SICA row and
    ' its current/historical HPS links; supplier enrichment belongs to a
    ' separate integration path and must not block local cache alignment.
    Set rsSource = p_SourceDb.OpenRecordset( _
        "SELECT TbUsuariosSICA.* FROM TbUsuariosSICA WHERE ID='" & SqlText(p_IDSICA) & "'", _
        dbOpenSnapshot)
    If rsSource.EOF Then GoTo SALIR

    Set rsTarget = p_LocalDb.OpenRecordset("SELECT * FROM " & p_TableName & " WHERE ID='" & SqlText(p_IDSICA) & "'", dbOpenDynaset)
    If rsTarget.EOF Then
        rsTarget.AddNew
    Else
        rsTarget.Edit
    End If

    For Each fld In rsTarget.Fields
        If FieldExists(rsSource, fld.Name) Then
            rsTarget.Fields(fld.Name).value = rsSource.Fields(fld.Name).value
        End If
    Next fld
    rsTarget.Update

SALIR:
    On Error Resume Next
    If Not rsTarget Is Nothing Then rsTarget.Close
    If Not rsSource Is Nothing Then rsSource.Close
    Set rsTarget = Nothing
    Set rsSource = Nothing
End Sub

Private Sub CopyUserFields(ByRef p_Source As DAO.Recordset, ByRef p_Target As DAO.Recordset)
    CopyFieldIfExists p_Source, p_Target, "DNI"
    CopyFieldIfExists p_Source, p_Target, "Nombre"
    CopyFieldIfExists p_Source, p_Target, "Apellido_1"
    CopyFieldIfExists p_Source, p_Target, "Apellido_2"
    CopyFieldIfExists p_Source, p_Target, "Telefono"
    CopyFieldIfExists p_Source, p_Target, "Correo_e"
    CopyFieldIfExists p_Source, p_Target, "F_Nacimiento"
    CopyFieldIfExists p_Source, p_Target, "Motivo_HPS"
    CopyFieldIfExists p_Source, p_Target, "IDExpediente"
    CopyFieldIfExists p_Source, p_Target, "F_Curso"
    CopyFieldIfExists p_Source, p_Target, "CursoEnVigor"
    CopyFieldIfExists p_Source, p_Target, "F_Baja"
    CopyFieldIfExists p_Source, p_Target, "FAvisoConcesion"
    CopyFieldIfExists p_Source, p_Target, "RequiereComunicacionConcesion"
End Sub

Private Sub ClearHpsFields(ByRef p_Target As DAO.Recordset)
    Dim prefixes As Variant
    Dim i As Long

    prefixes = Array("HPS_NAC", "HPS_OTAN", "HPS_ESA", "HPS_UE")
    For i = LBound(prefixes) To UBound(prefixes)
        SetFieldValue p_Target, CStr(prefixes(i)) & "_SIN_DATOS", "Sí"
        SetFieldValue p_Target, CStr(prefixes(i)) & "_F_Concesion", Null
        SetFieldValue p_Target, CStr(prefixes(i)) & "_F_Caducidad", Null
        SetFieldValue p_Target, CStr(prefixes(i)) & "_Grado", Null
        SetFieldValue p_Target, CStr(prefixes(i)) & "_F_Solicitud", Null
        SetFieldValue p_Target, CStr(prefixes(i)) & "_Especialidad", Null
        SetFieldValue p_Target, CStr(prefixes(i)) & "_Renovacion", Null
        SetFieldValue p_Target, CStr(prefixes(i)) & "_F_Baja", Null
    Next i
End Sub

Private Sub CopyHpsFields(ByRef p_Source As DAO.Recordset, ByRef p_Target As DAO.Recordset, ByVal p_Prefix As String)
    SetFieldValue p_Target, p_Prefix & "_SIN_DATOS", "No"
    CopyFieldIfExists p_Source, p_Target, "F_Concesion", p_Prefix & "_F_Concesion"
    CopyFieldIfExists p_Source, p_Target, "F_Caducidad", p_Prefix & "_F_Caducidad"
    CopyFieldIfExists p_Source, p_Target, "Grado", p_Prefix & "_Grado"
    CopyFieldIfExists p_Source, p_Target, "F_Solicitud", p_Prefix & "_F_Solicitud"
    CopyFieldIfExists p_Source, p_Target, "Especialidad", p_Prefix & "_Especialidad"
    CopyFieldIfExists p_Source, p_Target, "Renovacion", p_Prefix & "_Renovacion"
    CopyFieldIfExists p_Source, p_Target, "F_Baja", p_Prefix & "_F_Baja"
End Sub

Private Function HpsPrefix(ByVal p_TipoHPS As String) As String
    Select Case UCase$(Trim$(p_TipoHPS))
        Case "NACIONAL", "NAC"
            HpsPrefix = "HPS_NAC"
        Case "OTAN"
            HpsPrefix = "HPS_OTAN"
        Case "ESA"
            HpsPrefix = "HPS_ESA"
        Case "UE"
            HpsPrefix = "HPS_UE"
    End Select
End Function

Private Sub CopyFieldIfExists(ByRef p_Source As DAO.Recordset, ByRef p_Target As DAO.Recordset, ByVal p_SourceField As String, Optional ByVal p_TargetField As String = "")
    If p_TargetField = "" Then p_TargetField = p_SourceField
    If Not FieldExists(p_Source, p_SourceField) Then Exit Sub
    SetFieldValue p_Target, p_TargetField, p_Source.Fields(p_SourceField).value
End Sub

Private Sub SetFieldValue(ByRef p_Target As DAO.Recordset, ByVal p_FieldName As String, ByVal p_Value As Variant)
    On Error Resume Next
    p_Target.Fields(p_FieldName).value = p_Value
    Err.Clear
    On Error GoTo 0
End Sub

Private Function FieldExists(ByRef p_Recordset As DAO.Recordset, ByVal p_FieldName As String) As Boolean
    Dim tmp As String
    On Error Resume Next
    tmp = p_Recordset.Fields(p_FieldName).Name
    FieldExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function CountRowsInDb(ByRef p_Db As DAO.Database, ByVal p_TableName As String, ByVal p_WhereClause As String) As Long
    Dim rs As DAO.Recordset
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS C FROM " & p_TableName & " WHERE " & p_WhereClause, dbOpenSnapshot)
    CountRowsInDb = CLng(rs!C)
    rs.Close
    Set rs = Nothing
End Function

Private Function SqlText(ByVal value As String) As String
    SqlText = Replace(value, "'", "''")
End Function

