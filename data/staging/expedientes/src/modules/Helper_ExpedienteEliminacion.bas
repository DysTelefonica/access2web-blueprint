Attribute VB_Name = "Helper_ExpedienteEliminacion"
Option Compare Database
Option Explicit
' Helper REFAC-1c: eliminacion de expedientes.
' Stateless, DAO-injectable. Implementa los checks de MotivoEliminarNoOK
' (derivados, anexos, suministradores) como queries DAO puras, permitiendo
' tests deterministas.
'
' Cobertura: BR-26-01..03 (Verified-runtime), BR-26-04..05 (Verified-static, deferred).
'
' Tests en src/modules/Test_Helper_ExpedienteEliminacion.bas.
'
' PRUEBA-002 PR-C: agregar la cascade delete real (EliminarExpediente) que
' hoy vive como class method ExpedienteOperaciones.Eliminar.

' Resuelve la DAO: si el caller pasa una, usa esa; si no, getdb() del proyecto.
Private Function ResolveDb( _
    ByVal p_Db As DAO.Database, _
    ByRef p_Error As String) As DAO.Database
    On Error GoTo errores
    If p_Db Is Nothing Then
        Set ResolveDb = getdb()
    Else
        Set ResolveDb = p_Db
    End If
    Exit Function
errores:
    p_Error = "ResolveDb: " & Err.Description
End Function

' BR-26-01: Validacion - el expediente tiene hijos derivados (Lotes o Basados)?
' DADO un IDExpediente, CUANDO se consulta TbExpedientes WHERE IDExpedientePadre=p_ID
' ENTONCES retorna "NO" + motivo si N>0; "OK" si N=0; "ERR" si fallo DAO.
' DAO-injectable para tests deterministas.
Public Function TieneDerivados( _
    ByVal p_IDExpediente As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Motivo As String = "", _
    Optional ByRef p_Error As String) As String
    
    Dim dbUse As DAO.Database
    Dim rs As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    If p_IDExpediente <= 0 Then
        p_Error = "TieneDerivados: p_IDExpediente must be > 0 (got " & p_IDExpediente & ")"
        TieneDerivados = "ERR"
        Exit Function
    End If
    
    Set dbUse = ResolveDb(p_Db, p_Error)
    If dbUse Is Nothing Then
        TieneDerivados = "ERR"
        Exit Function
    End If
    
    m_SQL = "SELECT COUNT(*) AS N FROM TbExpedientes WHERE IDExpedientePadre=" & p_IDExpediente
    Set rs = dbUse.OpenRecordset(m_SQL, dbReadOnly)
    If rs!N > 0 Then
        p_Motivo = "De este elemento derivan otros expedientes (Lotes o Basados). Elimínelos primero."
        TieneDerivados = "NO"
    Else
        TieneDerivados = "OK"
    End If
    rs.Close
    
    Exit Function
errores:
    p_Error = "TieneDerivados: " & Err.Description
    TieneDerivados = "ERR"
End Function

' BR-26-02: Validacion - el expediente tiene anexos?
Public Function TieneAnexos( _
    ByVal p_IDExpediente As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Motivo As String = "", _
    Optional ByRef p_Error As String) As String
    
    Dim dbUse As DAO.Database
    Dim rs As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    If p_IDExpediente <= 0 Then
        p_Error = "TieneAnexos: p_IDExpediente must be > 0 (got " & p_IDExpediente & ")"
        TieneAnexos = "ERR"
        Exit Function
    End If
    
    Set dbUse = ResolveDb(p_Db, p_Error)
    If dbUse Is Nothing Then
        TieneAnexos = "ERR"
        Exit Function
    End If
    
    m_SQL = "SELECT COUNT(*) AS N FROM TbExpedientesAnexos WHERE IDExpediente=" & p_IDExpediente
    Set rs = dbUse.OpenRecordset(m_SQL, dbReadOnly)
    If rs!N > 0 Then
        p_Motivo = "Este expediente tiene Anexos. Elimínelos primero."
        TieneAnexos = "NO"
    Else
        TieneAnexos = "OK"
    End If
    rs.Close
    
    Exit Function
errores:
    p_Error = "TieneAnexos: " & Err.Description
    TieneAnexos = "ERR"
End Function

' BR-26-03: Validacion - el expediente tiene suministradores?
Public Function TieneSuministradores( _
    ByVal p_IDExpediente As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Motivo As String = "", _
    Optional ByRef p_Error As String) As String
    
    Dim dbUse As DAO.Database
    Dim rs As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    If p_IDExpediente <= 0 Then
        p_Error = "TieneSuministradores: p_IDExpediente must be > 0 (got " & p_IDExpediente & ")"
        TieneSuministradores = "ERR"
        Exit Function
    End If
    
    Set dbUse = ResolveDb(p_Db, p_Error)
    If dbUse Is Nothing Then
        TieneSuministradores = "ERR"
        Exit Function
    End If
    
    m_SQL = "SELECT COUNT(*) AS N FROM TbExpedientesSuministradores WHERE IDExpediente=" & p_IDExpediente
    Set rs = dbUse.OpenRecordset(m_SQL, dbReadOnly)
    If rs!N > 0 Then
        p_Motivo = "Este expediente tiene Suministradores. Elimínelos primero."
        TieneSuministradores = "NO"
    Else
        TieneSuministradores = "OK"
    End If
    rs.Close
    
    Exit Function
errores:
    p_Error = "TieneSuministradores: " & Err.Description
    TieneSuministradores = "ERR"
End Function

' BR-26-04: Wrapper de los 3 checks. Retorna "OK" si TODOS pasan; "NO" + motivo
' del primero que falle. Mantiene el contrato de la validacion previa de la
' ExpedienteOperaciones.Eliminar original.
'
' Futuros checks (post-PR-C): DPDs/Agedys (TbExpAgedys), GR (TbProyectosGestionRiesgos),
' HPS (TbExpedientesHPS), NCs (TbNoConformidades). Estos 4 son Verified-static por ahora.
Public Function PuedeEliminar( _
    ByVal p_IDExpediente As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Motivo As String = "", _
    Optional ByRef p_Error As String) As String
    
    Dim m_Resultado As String
    Dim m_MotivoLocal As String
    Dim m_ErrorLocal As String
    
    m_Resultado = TieneDerivados(p_IDExpediente, p_Db, m_MotivoLocal, m_ErrorLocal)
    If m_Resultado <> "OK" Then
        p_Motivo = m_MotivoLocal
        If m_Resultado = "ERR" Then p_Error = m_ErrorLocal
        PuedeEliminar = m_Resultado
        Exit Function
    End If
    
    m_Resultado = TieneAnexos(p_IDExpediente, p_Db, m_MotivoLocal, m_ErrorLocal)
    If m_Resultado <> "OK" Then
        p_Motivo = m_MotivoLocal
        If m_Resultado = "ERR" Then p_Error = m_ErrorLocal
        PuedeEliminar = m_Resultado
        Exit Function
    End If
    
    m_Resultado = TieneSuministradores(p_IDExpediente, p_Db, m_MotivoLocal, m_ErrorLocal)
    If m_Resultado <> "OK" Then
        p_Motivo = m_MotivoLocal
        If m_Resultado = "ERR" Then p_Error = m_ErrorLocal
        PuedeEliminar = m_Resultado
        Exit Function
    End If

    m_Resultado = TieneAgedys(p_IDExpediente, p_Db, m_MotivoLocal, m_ErrorLocal)
    If m_Resultado <> "OK" Then
        p_Motivo = m_MotivoLocal
        If m_Resultado = "ERR" Then p_Error = m_ErrorLocal
        PuedeEliminar = m_Resultado
        Exit Function
    End If

    PuedeEliminar = "OK"
End Function

' BR-26-09: Validacion - el expediente tiene DPDs (Agedys) vinculados?
' DADO un IDExpediente, CUANDO se consulta TbExpAgedys WHERE IdExpediente = p_IDExpediente
' ENTONCES retorna "NO" + motivo si N>0; "OK" si N=0; "ERR" si fallo DAO.
'
' Limitacion: 3 checks del original (GR via TbProyectosGestionRiesgos, HPS via TbHPS,
' NCs via TbNoConformidades) NO se implementan porque las tablas no existen en el backend
' de EXPEDIENTES - son de proyectos externos (GESTION_RIESGOS, NO_CONFORMIDADES, etc.).
' Estos 3 quedan como @Deprecated; la validacion real se hace en esos proyectos.
Public Function TieneAgedys( _
    ByVal p_IDExpediente As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Motivo As String = "", _
    Optional ByRef p_Error As String) As String
    
    Dim dbUse As DAO.Database
    Dim rs As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    If p_IDExpediente <= 0 Then
        p_Error = "TieneAgedys: p_IDExpediente must be > 0 (got " & p_IDExpediente & ")"
        TieneAgedys = "ERR"
        Exit Function
    End If
    
    Set dbUse = ResolveDb(p_Db, p_Error)
    If dbUse Is Nothing Then
        TieneAgedys = "ERR"
        Exit Function
    End If
    
    m_SQL = "SELECT COUNT(*) AS N FROM TbExpAgedys WHERE IdExpediente=" & p_IDExpediente
    Set rs = dbUse.OpenRecordset(m_SQL, dbReadOnly)
    If rs!N > 0 Then
        p_Motivo = "Este expediente está vinculado a los DPDs (Agedys). Elimínelos primero."
        TieneAgedys = "NO"
    Else
        TieneAgedys = "OK"
    End If
    rs.Close
    
    Exit Function
errores:
    p_Error = "TieneAgedys: " & Err.Description
    TieneAgedys = "ERR"
End Function

' BR-26-05: Cascade delete de un expediente y sus tablas hijas.
' Antes de eliminar, valida via PuedeEliminar. Si pasa, ejecuta el cascade en una
' transaccion (BeginTrans/CommitTrans) con rollback en error. DAO-injectable.
'
' Retorna:
'   "OK"  - eliminacion exitosa (cascade ejecutado, commit hecho)
'   "NO"  - validacion rechazo (p_Error con el motivo)
'   "ERR" - error de input o DAO (p_Error poblada)
'
' Orden del cascade (identico a ExpedienteOperaciones.Eliminar original):
'   A) TbExpedientesConEntidades (cache)
'   B) TbExpedientesSuministradores (tabla intermedia)
'   C) TbExpedientesHitos
'   D) TbExpedientesAnualidades
'   E) TbExpedientesModificados
'   F) TbExpedientesResponsables
'   G) Relaciones N:M (Comerciales, CPVs, Lugares, PECAL, RACS, CodigoCompras)
'   H) TbExpedientes (parent)
Public Function EliminarExpediente( _
    ByVal p_IDExpediente As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As String
    
    Dim dbUse As DAO.Database
    Dim ws As DAO.Workspace
    Dim m_TransaccionIniciadaAqui As Boolean
    Dim m_ResultadoValidacion As String
    Dim m_Motivo As String
    
    On Error GoTo errores
    
    ' Input validation
    If p_IDExpediente <= 0 Then
        p_Error = "EliminarExpediente: p_IDExpediente must be > 0 (got " & p_IDExpediente & ")"
        EliminarExpediente = "ERR"
        Exit Function
    End If
    
    ' Setup: si no nos pasan DB, iniciamos nuestra propia transaccion
    If p_Db Is Nothing Then
        Set dbUse = getdb()
        Set ws = DBEngine.Workspaces(0)
        ws.BeginTrans
        m_TransaccionIniciadaAqui = True
    Else
        ' Si nos pasan DB, nos unimos a la transaccion existente (no iniciamos la nuestra)
        Set dbUse = p_Db
        m_TransaccionIniciadaAqui = False
    End If
    
    ' Pre-validation
    m_ResultadoValidacion = PuedeEliminar(p_IDExpediente, dbUse, m_Motivo, p_Error)
    If m_ResultadoValidacion <> "OK" Then
        If m_ResultadoValidacion = "NO" Then p_Error = m_Motivo
        EliminarExpediente = m_ResultadoValidacion
        If m_TransaccionIniciadaAqui Then ws.Rollback
        Exit Function
    End If
    
    ' Cascade delete (orden inverso para integridad referencial)
    dbUse.Execute "DELETE * FROM TbExpedientesConEntidades WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientesSuministradores WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientesHitos WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientesAnualidades WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientesModificados WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientesResponsables WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientesComerciales WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientesCPVs WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientesLugaresEjecucion WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientesPECAL WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientesRACS WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientesCodigoCompras WHERE IDExpediente=" & p_IDExpediente
    dbUse.Execute "DELETE * FROM TbExpedientes WHERE IDExpediente=" & p_IDExpediente
    
    ' Commit
    If m_TransaccionIniciadaAqui Then
        ws.CommitTrans
    End If
    
    EliminarExpediente = "OK"
    Exit Function
    
errores:
    ' Rollback si iniciamos nosotros la transaccion
    If m_TransaccionIniciadaAqui And Not ws Is Nothing Then
        On Error Resume Next
        ws.Rollback
        On Error GoTo 0
    End If
    If Err.Number <> 0 And p_Error = "" Then
        p_Error = "EliminarExpediente: " & Err.Description
    End If
    EliminarExpediente = "ERR"
End Function
