Attribute VB_Name = "Mod_Sincronizacion_Historico"
Option Compare Database
Option Explicit

' =============================================================================================
' NOMBRE:       Mod_Sync_Historicos_Real
' DESCRIPCIÓN:  Sincronización Cross-Database (Remota -> Local).
'               Resuelve el problema de tablas no vinculadas y el fallo de INNER JOIN.
' =============================================================================================

''' <summary>
''' Regenera la tabla local TbUsuariosHistoricosLocal trayendo datos frescos del Backend (getdb).
''' </summary>

Public Function cache_usuarios_historicos_regenerar(Optional ByRef p_Error As String) As String

    Dim dbLocal As DAO.Database
    Dim m_SQL As String


    On Error GoTo errores

    Set dbLocal = CurrentDb()

    ' [PR1b] Incremental diff sync. Replaces the previous full-regen
    ' (DELETE all + INSERT all) with a per-row diff. Two statements:
    '   1. INSERT historical users from backend that are missing in
    '      local cache (handles c2h moves that left a row in backend
    '      without a matching local cache row).
    '   2. DELETE local cache rows for users that are no longer in
    '      the backend (handles h2c moves that left a stale row).
    ' Users that are in both with matching data are not touched.
    '
    ' Note: data drift in already-cached rows is NOT corrected here
    ' (e.g., a corrected Nombre in TbUsuariosHistoricos won't propagate
    ' until a full regen is explicitly triggered). Acceptable trade-off
    ' for the startup path; for a full regen the operator can call
    ' RegenerarUsuariosHistoricosLocalesFrontend.

    ' Step 1: wipe local cache. Follows the same DELETE-first -> INSERT-second pattern
    ' as Mod_Cache_Core.bas:cache_usuarios_regenerar1 (which deletes per-user before
    ' inserting). Bulk DELETE avoids the unique-index-on-DNI conflict that the
    ' diff-sync INSERT (with WHERE H.ID NOT IN) hits when the local table carries
    ' stale rows with the same DNI but a different ID.
    m_SQL = "DELETE FROM TbUsuariosHistoricosLocal"
    dbLocal.Execute m_SQL, dbFailOnError

    ' Step 2: insert every historical user fresh from the backend with the canonical
    ' 23-column explicit-INSERT pattern from src/queries/Consulta1.sql. No WHERE clause
    ' needed — the DELETE above guarantees an empty target.
    m_SQL = "INSERT INTO TbUsuariosHistoricosLocal ( EmpresaUsuario, EmpresaTramitadora, JuridicaContrato, CodExp, ID, DNI, Nombre, Apellido_1, Apellido_2, Telefono, Correo_e, IDExpediente, F_Nacimiento, LugarNacimiento, Motivo_HPS, F_Curso, CursoEnVigor, Requiere_Curso, FechaPrimeraConvocatoria, FechaSegundaConvocatoria, FechaCorreoNoCurso, F_Baja, LugarPrestacionServicio ) " & _
            "SELECT S1.Nombre AS EmpresaUsuario, S2.Nombre AS EmpresaTramitadora, H.CadenaContratistas AS JuridicaContrato, E.CodExp, H.ID, H.DNI, H.Nombre, H.Apellido_1, H.Apellido_2, H.Telefono, H.Correo_e, H.IDExpediente, H.F_Nacimiento, H.LugarNacimiento, H.Motivo_HPS, H.F_Curso, H.CursoEnVigor, H.Requiere_Curso, H.FechaPrimeraConvocatoria, H.FechaSegundaConvocatoria, H.FechaCorreoNoCurso, H.F_Baja, H.LugarPrestacionServicio " & _
            "FROM ((TbUsuariosHistoricos AS H LEFT JOIN TbSuministradores AS S1 ON H.IDEmpresaUsuario = S1.IDSuministrador) LEFT JOIN TbSuministradores AS S2 ON H.IDEmpresaHPS = S2.IDSuministrador) LEFT JOIN TbExpedientes AS E ON H.IDExpediente = E.IDExpediente"
    dbLocal.Execute m_SQL, dbFailOnError

    cache_usuarios_historicos_regenerar = "OK"


SALIR:
    On Error Resume Next

    Set dbLocal = Nothing

    Exit Function

errores:
    'ws.Rollback ' Deshacemos cambios si falla la red o el código
    p_Error = "Error Sync Históricos: " & Err.Description
    cache_usuarios_historicos_regenerar = "ERROR"
    Resume SALIR
End Function

''' <summary>
''' Regenera UN SOLO usuario histórico (Por ID).
''' Útil cuando acabas de pasar un usuario a histórico y quieres verlo ya.
''' </summary>
Public Function cache_usuario_historico_regenerar(p_ID As String, Optional ByRef p_Error As String) As String
    Dim dbRemota As DAO.Database
    Dim dbLocal As DAO.Database
    Dim rsOrigen As DAO.Recordset
    Dim rsDestino As DAO.Recordset
    Dim sSQL_Remota As String
    
    On Error GoTo errores
    
    Set dbRemota = getdb()
    Set dbLocal = CurrentDb()
    
    ' 1. Borrar versión vieja local
    dbLocal.Execute "DELETE FROM TbUsuariosHistoricosLocal WHERE ID=" & p_ID, dbFailOnError
    
    ' 2. Buscar datos frescos en Backend (LEFT JOIN para no perder datos)
    sSQL_Remota = "SELECT H.*, " & _
                  "S1.Nombre AS EmpresaUsuario, " & _
                  "S2.Nombre AS EmpresaTramitadora, " & _
                  "H.CadenaContratistas AS JuridicaContrato, " & _
                  "E.CodExp " & _
                  "FROM ((TbUsuariosHistoricos AS H " & _
                  "LEFT JOIN TbSuministradores AS S1 ON H.IDEmpresaUsuario = S1.IDSuministrador) " & _
                  "LEFT JOIN TbSuministradores AS S2 ON H.IDEmpresaHPS = S2.IDSuministrador) " & _
                  "LEFT JOIN TbExpedientes AS E ON H.IDExpediente = E.IDExpediente " & _
                  "WHERE H.ID = " & p_ID

    Set rsOrigen = dbRemota.OpenRecordset(sSQL_Remota, dbOpenSnapshot)
    
    If rsOrigen.EOF Then
        p_Error = "El ID " & p_ID & " no existe en la tabla histórica remota."
        cache_usuario_historico_regenerar = "NO_ENCONTRADO"
        GoTo SALIR
    End If
    
    ' 3. Insertar en Local
    Set rsDestino = dbLocal.OpenRecordset("SELECT * FROM TbUsuariosHistoricosLocal WHERE 1=0", dbOpenDynaset)
    rsDestino.AddNew
    
    Dim fld As DAO.Field
    For Each fld In rsDestino.Fields
        On Error Resume Next
        rsDestino(fld.Name).value = rsOrigen(fld.Name).value
        On Error GoTo errores
    Next fld
    rsDestino!ID = p_ID
    
    rsDestino.Update
    cache_usuario_historico_regenerar = "OK"

SALIR:
    On Error Resume Next
    rsOrigen.Close
    rsDestino.Close
    Exit Function
    
errores:
    p_Error = "Error Sync Individual: " & Err.Description
    cache_usuario_historico_regenerar = "ERROR"
End Function

''' <summary>
''' [PR1b cache] Atomic single-user cache synchronization after a lifecycle
''' move (c2h or h2c). Best-effort, does NOT raise on partial failure.
'''
''' Step 1: DELETE the user from ALL four cache tables (defensive cleanup
''' against partial state from a previous attempt that crashed mid-move).
''' Step 2: INSERT the user into the correct cache based on direction.
'''   - "c2h" -> historical cache (TbUsuariosHistoricosLocal)
'''   - "h2c" -> current cache (TbDatosLocal + TbUsuariosEntidades)
'''
''' Replaces what was previously a full regen of the current cache via
''' RefreshCaches. The full regen is still used at startup
''' (SincronizarCachesLocalesFrontend) where it makes sense; for the
''' single-user lifecycle move, atomic per-user ops are more efficient
''' AND cover both directions (c2h updates the historical cache, h2c
''' updates the current cache AND removes the user from the historical
''' cache, which the old behavior was leaving stale).
''' </summary>
Public Sub cache_usuario_sincronizar_despues_movimiento( _
    ByVal p_ID As String, _
    ByVal p_Direction As String, _
    Optional ByRef p_Error As String _
)
    Dim ws As DAO.Workspace
    Dim m_Err As String

    On Error GoTo errores

    p_Error = ""

    ' [PR2] Wrap the entire cache sync in a single transaction so partial
    ' state never leaks: if any of the 5 DELETEs or the directional INSERT
    ' fails, the cache rolls back atomically to its pre-sync state. The
    ' coordinator's CommitTrans has already happened (this runs post-commit),
    ' so a cache rollback leaves the move durable but the cache temporarily
    ' stale — the operator can trigger a manual regen to recover.
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans

    ' Step 1: DELETE the user from all five cache tables (defensive).
    ' Each On Error Resume Next block scopes a single DELETE so a failure
    ' on one table does not abort the others.
    On Error Resume Next
    CurrentDb().Execute "DELETE FROM TbDatosLocal WHERE ID=" & p_ID, dbFailOnError
    m_Err = Err.Description
    On Error GoTo errores
    If m_Err <> "" Then Debug.Print "[cache_usuario_sincronizar] DELETE TbDatosLocal: " & m_Err

    On Error Resume Next
    getdb().Execute "DELETE FROM TbUsuariosEntidades WHERE ID=" & p_ID, dbFailOnError
    m_Err = Err.Description
    On Error GoTo errores
    If m_Err <> "" Then Debug.Print "[cache_usuario_sincronizar] DELETE TbUsuariosEntidades: " & m_Err

    On Error Resume Next
    CurrentDb().Execute "DELETE FROM TbDatosLocalParaIndicadores WHERE ID=" & p_ID, dbFailOnError
    m_Err = Err.Description
    On Error GoTo errores
    If m_Err <> "" Then Debug.Print "[cache_usuario_sincronizar] DELETE TbDatosLocalParaIndicadores: " & m_Err

    ' [PR2] TbUsuariosSICALocalParaIndicadores.ID is TEXT (not numeric like
    ' TbDatosLocal.ID etc.), so the ID has to be single-quoted as a string in
    ' the WHERE clause — otherwise DAO 3464 ("data type mismatch in criteria
    ' expression") and the whole transaction rolls back.
    On Error Resume Next
    CurrentDb().Execute "DELETE FROM TbUsuariosSICALocalParaIndicadores WHERE ID='" & Replace(p_ID, "'", "''") & "'", dbFailOnError
    m_Err = Err.Description
    On Error GoTo errores
    If m_Err <> "" Then Debug.Print "[cache_usuario_sincronizar] DELETE TbUsuariosSICALocalParaIndicadores: " & m_Err

    On Error Resume Next
    CurrentDb().Execute "DELETE FROM TbUsuariosHistoricosLocal WHERE ID=" & p_ID, dbFailOnError
    m_Err = Err.Description
    On Error GoTo errores
    If m_Err <> "" Then Debug.Print "[cache_usuario_sincronizar] DELETE TbUsuariosHistoricosLocal: " & m_Err

    ' Step 2: INSERT into the correct cache based on direction.
    On Error Resume Next
    If p_Direction = "c2h" Then
        cache_usuario_historico_regenerar p_ID, p_Error
    ElseIf p_Direction = "h2c" Then
        cache_usuario_regenerar p_ID:=p_ID, p_Error:=p_Error
    End If
    On Error GoTo errores
    If p_Error <> "" Then
        Debug.Print "[cache_usuario_sincronizar] INSERT (" & p_Direction & "): " & p_Error
    End If

    ' [PR2] Commit the cache-sync transaction. If any error above caused
    ' Err to be set, we roll back. Defensive: if CommitTrans itself fails
    ' (e.g. transient JET issue), the EH handler will roll back too.
    On Error Resume Next
    ws.CommitTrans
    On Error GoTo errores

    Exit Sub

errores:
    ' [PR2] Roll back the cache-sync transaction on any error so the cache
    ' returns to its pre-sync state. Move has already been committed at
    ' the coordinator level, so the data is durable; the cache will be
    ' repaired by a manual regen.
    On Error Resume Next
    If Not ws Is Nothing Then ws.Rollback
    On Error GoTo 0
    p_Error = "cache_usuario_sincronizar_despues_movimiento: " & Err.Description
    Debug.Print "[cache_usuario_sincronizar] errores: " & p_Error
End Sub

