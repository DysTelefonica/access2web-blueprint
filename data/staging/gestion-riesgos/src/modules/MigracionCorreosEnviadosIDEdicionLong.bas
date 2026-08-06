Attribute VB_Name = "MigracionCorreosEnviadosIDEdicionLong"
Option Compare Database
Option Explicit

' ============================================================
' MigracionCorreosEnviadosIDEdicionLong -- issue #70
'
' Proposito: amplia TbCorreosEnviados.IDEdicion de Integer
'            (dbInteger, max 32,767) a Long (dbLong, max 2^31-1)
'            para que pueda almacenar valores de IDEdicion
'            provenientes de TbProyectosEdiciones.IDEdicion (Long).
'
' Comportamiento:
'   - Idempotente: si la columna ya es Long, retorna EnumSiNo.Si
'     sin re-ejecutar el ALTER.
'   - Pre-check: cuenta filas con IDEdicion > 32.767; si > 0, setea
'     p_Error con un mensaje descriptivo y retorna EnumSiNo.No
'     SIN propagar Err.Raise, para que el caller pueda reaccionar
'     programaticamente (respaldar, remediar, reintentar).
'   - Si el ALTER falla: retorna EnumSiNo.No y propaga
'     Err.Raise 1000 con source "MigracionCorreosEnviadosIDEdicionLong_<Helper>".
'
' Uso:
'   Dim errMsg As String
'   Dim resultado As EnumSiNo
'   resultado = MigracionCorreosEnviadosIDEdicionLong.Ejecutar(errMsg)
'   If errMsg <> "" Then ...  ' p_Error viaja por ByRef, nunca por Err
' ============================================================

' --- Module-level declarations (AGENTS.md rule 3) ---
Private Const INT_MAX_INCLUSIVE As Long = 32767   ' Limite superior real del tipo Integer DAO


' ============================================================
' Public entry point
' ============================================================
Public Function EjecutarMigracion(Optional ByRef p_Error As String) As EnumSiNo
    Dim db As DAO.Database

    On Error GoTo errores

    p_Error = ""

    ' Defense-in-depth: migration MUST run in testing mode (sandbox backend).
    ' If m_TestingMode=False here, something reset the state between test Setup
    ' and the EjecutarMigracion call (VBE cache, EVE, Application.Run context, etc.).
    ' Abort with a specific error instead of letting getdb fail with the opaque
    ' "m_ActiveBackendURL esta vacio" error from Variables Globales.bas:1069.
    If Not m_TestingMode Then
        p_Error = "EjecutarMigracion requires m_TestingMode=True (got False). " & _
                  "Likely causes: (1) test forgot to call EnsureTestConfigLoaded/ForceLocalBackend, " & _
                  "(2) VBE cache is stale -- recompile in Access VBE, " & _
                  "(3) EVE or ResetGlobals reset m_TestingMode after Setup, " & _
                  "(4) Application.Run lost the global context."
        Err.Raise 1000, "MigracionCorreosEnviadosIDEdicionLong_EjecutarMigracion", p_Error
    End If

    Set db = getdb(p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionCorreosEnviadosIDEdicionLong_EjecutarMigracion", p_Error
    If db Is Nothing Then
        p_Error = "No se pudo abrir la base de datos para ejecutar la migración"
        Err.Raise 1000, "MigracionCorreosEnviadosIDEdicionLong_EjecutarMigracion", p_Error
    End If

    ' 1. Idempotency guard: si la columna ya es Long, no-op
    Dim yaEsLong As Boolean
    yaEsLong = TipoColumnaEsLong_TbCorreosEnviados(db)
    If yaEsLong Then
        EjecutarMigracion = EnumSiNo.Sí
        Set db = Nothing
        Exit Function
    End If

    ' 2. Pre-check (filas con IDEdicion > Integer max)
    Dim countExcedidas As Long
    countExcedidas = PreCheck_TbCorreosEnviados_IDEdicion(db, p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionCorreosEnviadosIDEdicionLong_EjecutarMigracion", p_Error
    If countExcedidas > 0 Then
        ' Pre-check abort: retorno graceful (sin Err.Raise) para que el caller pueda
        ' respaldar la base, decidir qué hacer con esas filas y volver a correr la migración.
        p_Error = "ALTER abortado: " & countExcedidas & _
                  " fila(s) en TbCorreosEnviados tienen IDEdicion > " & INT_MAX_INCLUSIVE & _
                  ". Respaldar la base, remediar esos valores y volver a correr la migración."
        EjecutarMigracion = EnumSiNo.No
        Set db = Nothing
        Exit Function
    End If

    ' 3. ALTER COLUMN
    Dim alterResult As EnumSiNo
    alterResult = AlterTipoColumna_TbCorreosEnviados_IDEdicion(db, p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionCorreosEnviadosIDEdicionLong_EjecutarMigracion", p_Error
    If alterResult = EnumSiNo.No Then
        EjecutarMigracion = EnumSiNo.No
        Set db = Nothing
        Exit Function
    End If

    EjecutarMigracion = EnumSiNo.Sí
    Set db = Nothing
    Exit Function

errores:
    EjecutarMigracion = EnumSiNo.No
    If Err.Number <> 1000 Then
        p_Error = "El método EjecutarMigracion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Err.Raise Err.Number, "MigracionCorreosEnviadosIDEdicionLong_EjecutarMigracion", p_Error
End Function


' ============================================================
' Private helper: idempotency guard
' Retorna True si TbCorreosEnviados.IDEdicion ya es Long (dbLong=4).
' Retorna False si es Integer (dbInteger=3) o cualquier otro tipo.
' Retorna False ante cualquier error DAO (Err.Clear -- el orchestrator
' resuelve errores genuinos via el path ALTER).
' ============================================================
Private Function TipoColumnaEsLong_TbCorreosEnviados( _
    ByVal p_db As DAO.Database _
) As Boolean
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    On Error GoTo errores

    Set tdf = p_db.TableDefs("TbCorreosEnviados")
    Set fld = tdf.Fields("IDEdicion")
    TipoColumnaEsLong_TbCorreosEnviados = (fld.Type = dbLong)

    Set fld = Nothing
    Set tdf = Nothing
    Exit Function

errores:
    TipoColumnaEsLong_TbCorreosEnviados = False
    Err.Clear
End Function


' ============================================================
' Private helper: pre-check
' Retorna la cantidad de filas en TbCorreosEnviados con IDEdicion > 32767.
' Propaga errores DAO con Err.Raise 1000 (el orchestrator los maneja).
' ============================================================
Private Function PreCheck_TbCorreosEnviados_IDEdicion( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String _
) As Long
    Dim rs As DAO.Recordset

    On Error GoTo errores

    p_Error = ""
    Set rs = p_db.OpenRecordset( _
        "SELECT COUNT(*) AS C FROM TbCorreosEnviados WHERE IDEdicion > " & INT_MAX_INCLUSIVE, _
        dbOpenSnapshot)
    If Not rs.EOF Then
        PreCheck_TbCorreosEnviados_IDEdicion = CLng(Nz(rs.Fields("C").Value, 0))
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
    p_Error = "PreCheck_TbCorreosEnviados_IDEdicion: " & Err.Description
    Err.Raise 1000, "MigracionCorreosEnviadosIDEdicionLong_PreCheck", p_Error
End Function


' ============================================================
' Private helper: ALTER COLUMN IDEdicion LONG
' Retorna EnumSiNo.Si en exito, EnumSiNo.No ante fallo DAO (propaga Err.Raise 1000).
' ============================================================
Private Function AlterTipoColumna_TbCorreosEnviados_IDEdicion( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String _
) As EnumSiNo
    On Error GoTo errores

    p_Error = ""
    p_db.Execute "ALTER TABLE TbCorreosEnviados ALTER COLUMN IDEdicion LONG", dbFailOnError
    AlterTipoColumna_TbCorreosEnviados_IDEdicion = EnumSiNo.Sí
    Exit Function

errores:
    p_Error = "AlterTipoColumna_TbCorreosEnviados_IDEdicion: " & Err.Description
    AlterTipoColumna_TbCorreosEnviados_IDEdicion = EnumSiNo.No
    Err.Raise 1000, "MigracionCorreosEnviadosIDEdicionLong_AlterTipoColumna", p_Error
End Function