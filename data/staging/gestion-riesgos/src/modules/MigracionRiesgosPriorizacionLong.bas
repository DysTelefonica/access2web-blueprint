Attribute VB_Name = "MigracionRiesgosPriorizacionLong"
Option Compare Database
Option Explicit

' ============================================================
' MigracionRiesgosPriorizacionLong -- issue #106
'
' Proposito: amplia TbRiesgos.Priorizacion de Integer
'            (dbInteger, max 32,767) a Long (dbLong, max 2^31-1)
'            para evitar overflow cuando una edicion acumule mas de 32,767
'            posiciones de priorizacion.
'
' Comportamiento:
'   - Idempotente: si la columna ya es Long, retorna EnumSiNo.Si
'     sin re-ejecutar el ALTER.
'   - El entrypoint sin parametros es SOLO para validacion sandbox/test;
'     la ejecucion operacional en produccion se hace con SQL DBA del runbook.
'   - Pre-check: cuenta filas con Priorizacion > 32.767; si > 0, setea
'     p_Error con un mensaje descriptivo y retorna EnumSiNo.No
'     SIN propagar Err.Raise, para que el caller pueda reaccionar
'     programaticamente (respaldar, remediar, reintentar).
'   - Si el ALTER falla: retorna EnumSiNo.No y propaga
'     Err.Raise 1000 con source "MigracionRiesgosPriorizacionLong_<Helper>".
' ============================================================

' --- Module-level declarations (AGENTS.md rule 3) ---
Private Const INT_MAX_INCLUSIVE As Long = 32767   ' Limite superior real del tipo Integer DAO
Private Const TMP_PRIORIZACION_LONG As String = "Priorizacion_Issue106_Long"


' ============================================================
' Public entry point
' ============================================================
Public Function EjecutarMigracion(Optional ByRef p_Error As String) As EnumSiNo
    Dim db As DAO.Database
    Dim migrationResult As EnumSiNo

    On Error GoTo errores

    p_Error = ""

    ' Defense-in-depth: migration MUST run in testing mode (sandbox backend).
    ' Production execution is operational and must follow the runbook/DBA path.
    If Not m_TestingMode Then
        p_Error = "EjecutarMigracion requires m_TestingMode=True (got False). " & _
                  "Likely causes: (1) test forgot to call EnsureTestConfigLoaded/ForceLocalBackend, " & _
                  "(2) VBE cache is stale -- recompile in Access VBE, " & _
                  "(3) EVE or ResetGlobals reset m_TestingMode after Setup, " & _
                  "(4) Application.Run lost the global context."
        EjecutarMigracion = EnumSiNo.No
        Exit Function
    End If

    Set db = getdb(p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionRiesgosPriorizacionLong_EjecutarMigracion", p_Error
    If db Is Nothing Then
        p_Error = "No se pudo abrir la base de datos para ejecutar la migracion"
        Err.Raise 1000, "MigracionRiesgosPriorizacionLong_EjecutarMigracion", p_Error
    End If

    migrationResult = EjecutarMigracionEnBase(db, p_Error)
    EjecutarMigracion = migrationResult
    If migrationResult = EnumSiNo.No Then
        Set db = Nothing
        Exit Function
    End If
    If p_Error <> "" Then Err.Raise 1000, "MigracionRiesgosPriorizacionLong_EjecutarMigracion", p_Error

    Set db = Nothing
    Exit Function

errores:
    EjecutarMigracion = EnumSiNo.No
    If Err.Number <> 1000 Then
        p_Error = "El metodo EjecutarMigracion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Err.Raise Err.Number, "MigracionRiesgosPriorizacionLong_EjecutarMigracion", p_Error
End Function


' ============================================================
' Public helper: explicit database migration
' Usado por tests deterministas con temp backend controlado. En produccion,
' seguir el runbook DBA/manual SQL para backup, lock window y evidencia.
' ============================================================
Public Function EjecutarMigracionEnBase( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String _
) As EnumSiNo
    On Error GoTo errores

    p_Error = ""
    If p_db Is Nothing Then
        p_Error = "EjecutarMigracionEnBase requiere una DAO.Database abierta"
        Err.Raise 1000, "MigracionRiesgosPriorizacionLong_EjecutarMigracionEnBase", p_Error
    End If

    If Not EsBasePermitidaParaMigracion(p_db, p_Error) Then
        EjecutarMigracionEnBase = EnumSiNo.No
        Exit Function
    End If

    ' 1. Idempotency guard: si la columna ya es Long, no-op
    Dim yaEsLong As Boolean
    yaEsLong = TipoColumnaEsLong_TbRiesgos(p_db)
    If yaEsLong Then
        EjecutarMigracionEnBase = EnumSiNo.Sí
        Exit Function
    End If

    ' 2. Pre-check (filas con Priorizacion > Integer max)
    Dim countExcedidas As Long
    countExcedidas = PreCheck_TbRiesgos_Priorizacion(p_db, p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionRiesgosPriorizacionLong_EjecutarMigracionEnBase", p_Error
    If countExcedidas > 0 Then
        ' Pre-check abort: retorno graceful (sin Err.Raise) para que el caller pueda
        ' respaldar la base, decidir que hacer con esas filas y volver a correr la migracion.
        p_Error = "ALTER abortado: " & countExcedidas & _
                  " fila(s) en TbRiesgos tienen Priorizacion > " & INT_MAX_INCLUSIVE & _
                  ". Respaldar la base, remediar esos valores y volver a correr la migracion."
        EjecutarMigracionEnBase = EnumSiNo.No
        Exit Function
    End If

    ' 3. ALTER COLUMN
    Dim alterResult As EnumSiNo
    alterResult = AlterTipoColumna_TbRiesgos_Priorizacion(p_db, p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionRiesgosPriorizacionLong_EjecutarMigracionEnBase", p_Error
    If alterResult = EnumSiNo.No Then
        EjecutarMigracionEnBase = EnumSiNo.No
        Exit Function
    End If

    EjecutarMigracionEnBase = EnumSiNo.Sí
    Exit Function

errores:
    EjecutarMigracionEnBase = EnumSiNo.No
    If Err.Number <> 1000 Then
        p_Error = "El metodo EjecutarMigracionEnBase ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Err.Raise Err.Number, "MigracionRiesgosPriorizacionLong_EjecutarMigracionEnBase", p_Error
End Function


' ============================================================
' Private helper: explicit DB safety guard
' Allows the normal testing-mode sandbox path and the isolated temp
' backend used by the cold-schema test. Blocks arbitrary production DBs.
' ============================================================
Private Function EsBasePermitidaParaMigracion( _
    ByVal p_db As DAO.Database, _
    ByRef p_Error As String _
) As Boolean
    Dim dbPath As String
    Dim tempPath As String
    Dim tmpPath As String
    Dim dbPathLower As String
    Dim allowedTempPrefix As String

    On Error GoTo errores

    If m_TestingMode Then
        EsBasePermitidaParaMigracion = True
        Exit Function
    End If

    dbPath = p_db.Name
    dbPathLower = LCase$(dbPath)
    tempPath = LCase$(Environ$("TEMP"))
    tmpPath = LCase$(Environ$("TMP"))
    If tempPath <> "" And Right$(tempPath, 1) <> "\" Then tempPath = tempPath & "\"
    If tmpPath <> "" And Right$(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
    allowedTempPrefix = LCase$("GestionRiesgos_Issue106_")

    If dbPathLower <> "" Then
        If (tempPath <> "" And Left$(dbPathLower, Len(tempPath)) = tempPath) Or _
           (tmpPath <> "" And Left$(dbPathLower, Len(tmpPath)) = tmpPath) Then
            If InStr(1, LCase$(Dir$(dbPath)), allowedTempPrefix, vbTextCompare) = 1 Then
                EsBasePermitidaParaMigracion = True
                Exit Function
            End If
        End If
    End If

    p_Error = "EjecutarMigracionEnBase bloqueado: la base explícita no parece sandbox/testing ni temp backend controlado issue #106. Ruta: " & dbPath
    EsBasePermitidaParaMigracion = False
    Exit Function

errores:
    p_Error = "EsBasePermitidaParaMigracion: " & Err.Description
    EsBasePermitidaParaMigracion = False
End Function


' ============================================================
' Private helper: idempotency guard
' Retorna True si TbRiesgos.Priorizacion ya es Long (dbLong=4).
' Retorna False si es Integer (dbInteger=3) o cualquier otro tipo.
' Retorna False ante cualquier error DAO (Err.Clear -- el orchestrator
' resuelve errores genuinos via el path ALTER).
' ============================================================
Private Function TipoColumnaEsLong_TbRiesgos( _
    ByVal p_db As DAO.Database _
) As Boolean
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    On Error GoTo errores

    Set tdf = p_db.TableDefs("TbRiesgos")
    Set fld = tdf.Fields("Priorizacion")
    TipoColumnaEsLong_TbRiesgos = (fld.Type = dbLong)

    Set fld = Nothing
    Set tdf = Nothing
    Exit Function

errores:
    TipoColumnaEsLong_TbRiesgos = False
    Err.Clear
End Function


' ============================================================
' Private helper: pre-check
' Retorna la cantidad de filas en TbRiesgos con Priorizacion > 32767.
' Propaga errores DAO con Err.Raise 1000 (el orchestrator los maneja).
' ============================================================
Private Function PreCheck_TbRiesgos_Priorizacion( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String _
) As Long
    Dim rs As DAO.Recordset

    On Error GoTo errores

    p_Error = ""
    Set rs = p_db.OpenRecordset( _
        "SELECT COUNT(*) AS C FROM TbRiesgos WHERE Priorizacion > " & INT_MAX_INCLUSIVE, _
        dbOpenSnapshot)
    If Not rs.EOF Then
        PreCheck_TbRiesgos_Priorizacion = CLng(Nz(rs.Fields("C").Value, 0))
    End If
    rs.Close
    Set rs = Nothing
    Exit Function

errores:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0
    p_Error = "PreCheck_TbRiesgos_Priorizacion: " & Err.Description
    Err.Raise 1000, "MigracionRiesgosPriorizacionLong_PreCheck", p_Error
End Function


' ============================================================
' Private helper: ALTER COLUMN Priorizacion LONG
' Retorna EnumSiNo.Si en exito, EnumSiNo.No ante fallo DAO (propaga Err.Raise 1000).
' ============================================================
Private Function AlterTipoColumna_TbRiesgos_Priorizacion( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String _
) As EnumSiNo
    On Error GoTo errores

    p_Error = ""
    p_db.Execute "ALTER TABLE TbRiesgos ALTER COLUMN Priorizacion LONG", dbFailOnError
    p_db.TableDefs.Refresh
    If TipoColumnaEsLong_TbRiesgos(p_db) Then
        AlterTipoColumna_TbRiesgos_Priorizacion = EnumSiNo.Sí
        Exit Function
    End If

    AlterTipoColumna_TbRiesgos_Priorizacion = ReconstruirColumnaPriorizacionLong(p_db, p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionRiesgosPriorizacionLong_AlterTipoColumna", p_Error
    If AlterTipoColumna_TbRiesgos_Priorizacion = EnumSiNo.No Then Exit Function

    If Not TipoColumnaEsLong_TbRiesgos(p_db) Then
        p_Error = "ALTER/Fallback no dejo TbRiesgos.Priorizacion como dbLong"
        Err.Raise 1000, "MigracionRiesgosPriorizacionLong_AlterTipoColumna", p_Error
    End If

    AlterTipoColumna_TbRiesgos_Priorizacion = EnumSiNo.Sí
    Exit Function

errores:
    p_Error = "AlterTipoColumna_TbRiesgos_Priorizacion: " & Err.Description
    AlterTipoColumna_TbRiesgos_Priorizacion = EnumSiNo.No
    Err.Raise 1000, "MigracionRiesgosPriorizacionLong_AlterTipoColumna", p_Error
End Function


' ============================================================
' Private helper: DAO fallback when Access SQL ALTER COLUMN LONG
' completes but leaves a DAO-created dbInteger field as dbInteger.
' Preserves existing values by copying through a temporary dbLong field.
' ============================================================
Private Function ReconstruirColumnaPriorizacionLong( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String _
) As EnumSiNo
    Dim tdf As DAO.TableDef
    Dim tmpField As DAO.Field
    Dim newField As DAO.Field

    On Error GoTo errores

    p_Error = ""
    Set tdf = p_db.TableDefs("TbRiesgos")

    On Error Resume Next
    tdf.Fields.Delete TMP_PRIORIZACION_LONG
    Err.Clear
    On Error GoTo errores

    Set tmpField = tdf.CreateField(TMP_PRIORIZACION_LONG, dbLong)
    tdf.Fields.Append tmpField
    p_db.TableDefs.Refresh

    p_db.Execute "UPDATE TbRiesgos SET " & TMP_PRIORIZACION_LONG & " = Priorizacion", dbFailOnError

    tdf.Fields.Delete "Priorizacion"
    p_db.TableDefs.Refresh

    Set tdf = p_db.TableDefs("TbRiesgos")
    Set newField = tdf.CreateField("Priorizacion", dbLong)
    tdf.Fields.Append newField
    p_db.TableDefs.Refresh

    p_db.Execute "UPDATE TbRiesgos SET Priorizacion = " & TMP_PRIORIZACION_LONG, dbFailOnError

    Set tdf = p_db.TableDefs("TbRiesgos")
    tdf.Fields.Delete TMP_PRIORIZACION_LONG
    p_db.TableDefs.Refresh

    ReconstruirColumnaPriorizacionLong = EnumSiNo.Sí
    Set newField = Nothing
    Set tmpField = Nothing
    Set tdf = Nothing
    Exit Function

errores:
    p_Error = "ReconstruirColumnaPriorizacionLong: " & Err.Description
    ReconstruirColumnaPriorizacionLong = EnumSiNo.No
    Err.Raise 1000, "MigracionRiesgosPriorizacionLong_ReconstruirColumna", p_Error
End Function
