Attribute VB_Name = "MigracionTbCambiosParaPublicacionEdicionLong"
Option Compare Database
Option Explicit

' ============================================================
' MigracionTbCambiosParaPublicacionEdicionLong -- issue #107
'
' Proposito: amplia tbCambiosParaPublicacion.EdicionInicial y
'            tbCambiosParaPublicacion.EdicionFinal de Integer
'            (dbInteger, max 32,767) a Long (dbLong, max 2^31-1)
'            para que puedan almacenar valores de IDEdicion
'            provenientes de TbProyectosEdiciones.IDEdicion (Long).
'
' Comportamiento:
'   - Idempotente: si AMBAS columnas ya son Long, retorna
'     EnumSiNo.Si sin re-ejecutar el ALTER.
'   - Pre-check: cuenta filas con EdicionInicial > 32.767 o
'     EdicionFinal > 32.767; si > 0 para cualquiera, setea
'     p_Error con un mensaje descriptivo y retorna EnumSiNo.No
'     SIN propagar Err.Raise, para que el caller pueda reaccionar
'     programaticamente (respaldar, remediar, reintentar).
'   - Si el ALTER falla: retorna EnumSiNo.No y propaga
'     Err.Raise 1000 con source "MigracionTbCambiosParaPublicacionEdicionLong_<Helper>".
'
' A diferencia de MigracionCorreosEnviadosIDEdicionLong (#70),
' ESTE modulo migra DOS columnas. La idempotencia se evalua sobre
' AMBAS: si solo EdicionInicial es Long pero EdicionFinal sigue
' siendo Integer, el helper lo detecta y ALTERA la restante,
' retornando EnumSiNo.Si al final. Esto cubre el caso de una
' migracion parcial previa abortada.
'
' Uso:
'   Dim errMsg As String
'   Dim resultado As EnumSiNo
'   resultado = MigracionTbCambiosParaPublicacionEdicionLong.EjecutarMigracion(errMsg)
'   If errMsg <> "" Then ...  ' p_Error viaja por ByRef, nunca por Err
' ============================================================

' --- Module-level declarations (AGENTS.md rule 3) ---
Private Const INT_MAX_INCLUSIVE As Long = 32767   ' Limite superior real del tipo Integer DAO
Private Const TABLE_NAME As String = "tbCambiosParaPublicacion"


' ============================================================
' Public entry point
' ============================================================
Public Function EjecutarMigracion(Optional ByRef p_Error As String) As EnumSiNo
    Dim db As DAO.Database

    On Error GoTo errores

    p_Error = ""

    ' Defense-in-depth: migration MUST run in testing mode (sandbox backend).
    ' Same rationale as #70 (lines 44-56 of MigracionCorreosEnviadosIDEdicionLong.bas).
    If Not m_TestingMode Then
        p_Error = "EjecutarMigracion requires m_TestingMode=True (got False). " & _
                  "Likely causes: (1) test forgot to call EnsureTestConfigLoaded/ForceLocalBackend, " & _
                  "(2) VBE cache is stale -- recompile in Access VBE, " & _
                  "(3) EVE or ResetGlobals reset m_TestingMode after Setup, " & _
                  "(4) Application.Run lost the global context."
        Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_EjecutarMigracion", p_Error
    End If

    Set db = getdb(p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_EjecutarMigracion", p_Error
    If db Is Nothing Then
        p_Error = "No se pudo abrir la base de datos para ejecutar la migración"
        Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_EjecutarMigracion", p_Error
    End If

    ' 1. Idempotency guard: si AMBAS columnas ya son Long, no-op
    Dim estadoActual As String
    estadoActual = EstadoColumnas_tbCambiosParaPublicacion(db, p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_EjecutarMigracion", p_Error
    ' Defense-in-depth (lesson 2026-07-14 issue #121): si EstadoColumnas devuelve "" por cualquier error DAO silencioso,
    ' NO retornamos Si. Si no, la migracion seria un silent-no-op y el caller creeria que se ejecuto.
    If Not IsEstadoColumnasValido(estadoActual) Then
        p_Error = "EstadoColumnas devolvio un estado no reconocible: '" & estadoActual & _
                  "'. No se puede determinar si la migracion es necesaria. " & _
                  "Posible causa: tabla tbCambiosParaPublicacion no existe en el backend actual, " & _
                  "o los nombres de campo cambiaron."
        Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_EjecutarMigracion", p_Error
    End If
    If estadoActual = "LONG,LONG" Then
        EjecutarMigracion = EnumSiNo.Sí
        Set db = Nothing
        Exit Function
    End If

    ' 2. Pre-check (filas con EdicionInicial o EdicionFinal > Integer max)
    Dim countExcedidas As Long
    countExcedidas = PreCheck_tbCambiosParaPublicacion(db, p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_EjecutarMigracion", p_Error
    If countExcedidas > 0 Then
        ' Pre-check abort: retorno graceful (sin Err.Raise) para que el caller pueda
        ' respaldar la base, decidir qué hacer con esas filas y volver a correr la migración.
        p_Error = "ALTER abortado: " & countExcedidas & _
                  " fila(s) en tbCambiosParaPublicacion tienen EdicionInicial o EdicionFinal > " & INT_MAX_INCLUSIVE & _
                  ". Respaldar la base, remediar esos valores y volver a correr la migración."
        EjecutarMigracion = EnumSiNo.No
        Set db = Nothing
        Exit Function
    End If

    ' 3. ALTER COLUMN — ambas, en orden. Cada ALTER valida pre/post tipo para ese campo especifico.
    '    Si estadoActual = "INTEGER,LONG", alteramos solo EdicionInicial.
    '    Si estadoActual = "LONG,INTEGER", alteramos solo EdicionFinal.
    '    Si estadoActual = "INTEGER,INTEGER", alteramos ambas (orden arbitrario; ambas son additivas).
    Dim alterResult As EnumSiNo
    If Left$(estadoActual, 6) = "INTEGER" Then
        alterResult = AlterTipoColumna_tbCambiosParaPublicacion_EdicionInicial(db, p_Error)
        If p_Error <> "" Then Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_EjecutarMigracion", p_Error
        If alterResult = EnumSiNo.No Then
            EjecutarMigracion = EnumSiNo.No
            Set db = Nothing
            Exit Function
        End If
    End If
    If Right$(estadoActual, 6) = "INTEGER" Then
        alterResult = AlterTipoColumna_tbCambiosParaPublicacion_EdicionFinal(db, p_Error)
        If p_Error <> "" Then Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_EjecutarMigracion", p_Error
        If alterResult = EnumSiNo.No Then
            EjecutarMigracion = EnumSiNo.No
            Set db = Nothing
            Exit Function
        End If
    End If

    ' 4. Defense-in-depth post-ALTER: re-leer tipos y verificar que efectivamente quedaron LONG.
    '    Si los ALTER fallaron silenciosamente (p.ej. permisos, tabla en diseno por otra sesion Access),
    '    esta verificacion los caza y devuelve No con p_Error descriptivo.
    Dim estadoPost As String
    estadoPost = EstadoColumnas_tbCambiosParaPublicacion(db, p_Error)
    If p_Error <> "" Then
        EjecutarMigracion = EnumSiNo.No
        Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_EjecutarMigracion", _
                  "Post-ALTER EstadoColumnas fallo: " & p_Error
    End If
    If estadoPost = "LONG,LONG" Then
        EjecutarMigracion = EnumSiNo.Sí
    Else
        p_Error = "ALTER no converio. Tipo esperado LONG,LONG; tipo actual " & estadoPost & _
                  ". Revisar permisos del backend, que no haya FK/declares bloqueando el cambio de tipo, " & _
                  "y que no haya otra sesion Access abierta bloqueando la tabla en modo diseno."
        EjecutarMigracion = EnumSiNo.No
    End If
    Set db = Nothing
    Exit Function

errores:
    EjecutarMigracion = EnumSiNo.No
    If Err.Number <> 1000 Then
        p_Error = "El método EjecutarMigracion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Err.Raise Err.Number, "MigracionTbCambiosParaPublicacionEdicionLong_EjecutarMigracion", p_Error
End Function


' ============================================================
' Private helper: valida que un estado devuelto por EstadoColumnas
' tenga el formato esperado ("INTEGER,INTEGER" / "INTEGER,LONG"
' / "LONG,INTEGER" / "LONG,LONG"). Si no, NO retornamos Si del caller
' para evitar silent-no-op (lección del 2026-07-14 issue #121).
' Acepta solo las 4 combinaciones canonicas. Si TipoComoNombre evoluciona,
' anadir las nuevas combos a este set explicito.
' ============================================================
Private Function IsEstadoColumnasValido(ByVal p_estado As String) As Boolean
    Dim estadosValidos As Variant
    estadosValidos = Array("INTEGER,INTEGER", "INTEGER,LONG", "LONG,INTEGER", "LONG,LONG")
    Dim i As Long
    For i = LBound(estadosValidos) To UBound(estadosValidos)
        If StrComp(p_estado, CStr(estadosValidos(i)), vbTextCompare) = 0 Then
            IsEstadoColumnasValido = True
            Exit Function
        End If
    Next i
    IsEstadoColumnasValido = False
End Function


' ============================================================
' Private helper: idempotency guard (devuelve "<tipoIni>,<tipoFin>")
'   "INTEGER,LONG"  — solo EdicionFinal migrada (migracion parcial previa)
'   "LONG,INTEGER"  — solo EdicionInicial migrada (caso simetrico)
'   "LONG,LONG"     — ambas migradas (idempotent no-op)
'   "INTEGER,INTEGER" — ninguna migrada (caso fresh)
' Retorna "" ante cualquier error DAO (Err.Clear -- el orchestrator
' resuelve errores genuinos via el path ALTER).
' ============================================================
Private Function EstadoColumnas_tbCambiosParaPublicacion( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String _
) As String
    Dim tdf As DAO.TableDef
    Dim fldIni As DAO.Field
    Dim fldFin As DAO.Field

    On Error GoTo errores

    Set tdf = p_db.TableDefs(TABLE_NAME)
    Set fldIni = tdf.Fields("EdicionInicial")
    Set fldFin = tdf.Fields("EdicionFinal")
    EstadoColumnas_tbCambiosParaPublicacion = TipoComoNombre(fldIni.Type) & "," & TipoComoNombre(fldFin.Type)

    Set fldIni = Nothing
    Set fldFin = Nothing
    Set tdf = Nothing
    Exit Function

errores:
    EstadoColumnas_tbCambiosParaPublicacion = ""
    Err.Clear
End Function


' ============================================================
' Private helper: convierte DAO dbType a nombre legible
' ============================================================
Private Function TipoComoNombre(ByVal dbType As Long) As String
    Select Case dbType
        Case dbLong
            TipoComoNombre = "LONG"
        Case dbInteger
            TipoComoNombre = "INTEGER"
        Case Else
            TipoComoNombre = "OTHER"
    End Select
End Function


' ============================================================
' Private helper: pre-check
' Retorna la cantidad de filas en tbCambiosParaPublicacion con
' EdicionInicial > 32767 OR EdicionFinal > 32767.
' Propaga errores DAO con Err.Raise 1000 (el orchestrator los maneja).
'
' Por que cubre las DOS columnas: el caller (EjecutarMigracion)
' necesita abortar si CUALQUIERA de las dos tiene valores fuera
' de rango -- el ALTER falla irreversiblemente si encuentra
' datos truncables.
' ============================================================
Private Function PreCheck_tbCambiosParaPublicacion( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String _
) As Long
    Dim rs As DAO.Recordset

    On Error GoTo errores

    p_Error = ""
    Set rs = p_db.OpenRecordset( _
        "SELECT COUNT(*) AS C FROM tbCambiosParaPublicacion " & _
        "WHERE EdicionInicial > " & INT_MAX_INCLUSIVE & _
        " OR EdicionFinal > " & INT_MAX_INCLUSIVE, _
        dbOpenSnapshot)
    If Not rs.EOF Then
        PreCheck_tbCambiosParaPublicacion = CLng(Nz(rs.Fields("C").Value, 0))
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
    p_Error = "PreCheck_tbCambiosParaPublicacion: " & Err.Description
    Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_PreCheck", p_Error
End Function


' ============================================================
' Private helper: ALTER COLUMN EdicionInicial LONG
' Retorna EnumSiNo.Si en exito, EnumSiNo.No ante fallo DAO (propaga Err.Raise 1000).
' ============================================================
Private Function AlterTipoColumna_tbCambiosParaPublicacion_EdicionInicial( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String _
) As EnumSiNo
    On Error GoTo errores

    p_Error = ""
    p_db.Execute "ALTER TABLE tbCambiosParaPublicacion ALTER COLUMN EdicionInicial LONG", dbFailOnError
    AlterTipoColumna_tbCambiosParaPublicacion_EdicionInicial = EnumSiNo.Sí
    Exit Function

errores:
    p_Error = "AlterTipoColumna_tbCambiosParaPublicacion_EdicionInicial: " & Err.Description
    AlterTipoColumna_tbCambiosParaPublicacion_EdicionInicial = EnumSiNo.No
    Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_AlterTipoColumna_EdicionInicial", p_Error
End Function


' ============================================================
' Private helper: ALTER COLUMN EdicionFinal LONG
' Retorna EnumSiNo.Si en exito, EnumSiNo.No ante fallo DAO (propaga Err.Raise 1000).
' ============================================================
Private Function AlterTipoColumna_tbCambiosParaPublicacion_EdicionFinal( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String _
) As EnumSiNo
    On Error GoTo errores

    p_Error = ""
    p_db.Execute "ALTER TABLE tbCambiosParaPublicacion ALTER COLUMN EdicionFinal LONG", dbFailOnError
    AlterTipoColumna_tbCambiosParaPublicacion_EdicionFinal = EnumSiNo.Sí
    Exit Function

errores:
    p_Error = "AlterTipoColumna_tbCambiosParaPublicacion_EdicionFinal: " & Err.Description
    AlterTipoColumna_tbCambiosParaPublicacion_EdicionFinal = EnumSiNo.No
    Err.Raise 1000, "MigracionTbCambiosParaPublicacionEdicionLong_AlterTipoColumna_EdicionFinal", p_Error
End Function
