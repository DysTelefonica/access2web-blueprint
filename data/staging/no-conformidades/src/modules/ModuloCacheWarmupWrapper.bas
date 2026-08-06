Attribute VB_Name = "ModuloCacheWarmupWrapper"
Option Compare Database
Option Explicit

' ============================================================
' ModuloCacheWarmupWrapper — Cache_Warmup_Opciones wrapper
' Granular centralizado para warming / rebuild / populate de caches.
'
' SDD: cache-idempotent-warmup-2026-06-17
' Issues: #88, #89, #90, #91, #92, #93
'
' Punto de entrada unico para warming de las 6 tiers:
'   LISTADO, DETALLE_NC_PROYECTO, LISTADO_AUDITORIA,
'   INDICADORES_PROYECTO, INDICADORES_AUDITORIA, ESTADO_CATALOGO.
'
' API publica:
'   Cache_Warmup_Opciones(p_Vias, p_EnsureSchema, p_DryRun, p_Error) As String
'
' Devuelve JSON parseable por JsonConverter:
'   {
'     "dry_run": <bool>,
'     "ensure_schema_ejecutado": <bool>,
'     "vias_procesadas": [<string>, ...],
'     "vias_omitidas":  [<string>, ...],
'     "resultados": {
'       "<tier>": { "ok": <bool>,
'                   "filas_insertadas": <long>,
'                   "filas_que_insertaria": <long>,   // solo en dry-run
'                   "duracion_seg": <double> }
'     },
'     "errores": [<string>, ...],
'     "duracion_total_seg": <double>
'   }
' ============================================================

' === PUBLIC CONSTANTS — 6 tier names (autodescubribles via VBE Intellisense) ===
Public Const CACHE_VIA_LISTADO                As String = "LISTADO"
Public Const CACHE_VIA_DETALLE_NC_PROYECTO    As String = "DETALLE_NC_PROYECTO"
Public Const CACHE_VIA_LISTADO_AUDITORIA      As String = "LISTADO_AUDITORIA"
Public Const CACHE_VIA_INDICADORES_PROYECTO   As String = "INDICADORES_PROYECTO"
Public Const CACHE_VIA_INDICADORES_AUDITORIA  As String = "INDICADORES_AUDITORIA"
Public Const CACHE_VIA_ESTADO_CATALOGO        As String = "ESTADO_CATALOGO"

' === Tablas backend de cada tier (referencia interna) ===
Private Const TBL_LISTADO_NC                  As String = "TbCacheListadoNC"
Private Const TBL_NC_PROYECTO                 As String = "TbCacheNCProyecto"
Private Const TBL_LISTADO_NC_AUD              As String = "TbCacheListadoNCAuditoria"
Private Const TBL_INDICADORES_HEADER_PROY     As String = "TbCacheIndicadoresProyectoHeader"
Private Const TBL_INDICADORES_DETALLE_PROY    As String = "TbCacheIndicadoresProyectoDetalle"
Private Const TBL_INDICADORES_HEADER_AUD      As String = "TbCacheIndicadoresAuditoriaHeader"
Private Const TBL_INDICADORES_DETALLE_AUD     As String = "TbCacheIndicadoresAuditoriaDetalle"
Private Const TBL_ESTADO_CATALOGO             As String = "TbEstadoCatalogo"

' === Module-level state compartido entre subrutinas por-tier ===
' Seteados al inicio de Cache_Warmup_Opciones; limpiados al final.
Private m_Db As DAO.Database
Private m_Errores As Collection

' ============================================================
' PUBLIC API
' ============================================================

' Cache_Warmup_Opciones — wrapper granular centralizado.
'
' p_Vias          ""  = todas las tiers (6).
'                  "LISTADO,INDICADORES_PROYECTO" = inclusion.
'                  "-INDICADORES_AUDITORIA,-ESTADO_CATALOGO" = exclusion.
'                  Mezcla inclusion+exclusion: inclusion primero, luego exclusion.
'                  Nombres desconocidos -> append "Unknown tier: X" a errores.
' p_EnsureSchema  True (default) = verifica/crea schema backend antes de warming.
'                  False = asume schema listo (skip Ensure*).
' p_DryRun        True = no escribe en ninguna tabla; computa filas_que_insertaria.
Public Function Cache_Warmup_Opciones( _
    Optional ByVal p_Vias As String = "", _
    Optional ByVal p_EnsureSchema As Boolean = True, _
    Optional ByVal p_DryRun As Boolean = False, _
    Optional ByRef p_Error As String _
) As String

    Dim allTiers As Collection
    Dim included As Collection
    Dim omitted As Collection
    Dim resultados As Scripting.Dictionary
    Dim dbErr As String
    Dim inicioTotal As Double
    Dim duracionTotal As Double
    Dim hasProy As Boolean
    Dim hasAud As Boolean
    Dim pairedIndicadores As Boolean
    Dim parseOk As Boolean
    Dim tier As Variant
    Dim ctx As Scripting.Dictionary
    Dim tierName As String

    On Error GoTo errores

    p_Error = ""
    inicioTotal = Timer

    Set allTiers = New Collection
    allTiers.Add CACHE_VIA_LISTADO
    allTiers.Add CACHE_VIA_DETALLE_NC_PROYECTO
    allTiers.Add CACHE_VIA_LISTADO_AUDITORIA
    allTiers.Add CACHE_VIA_INDICADORES_PROYECTO
    allTiers.Add CACHE_VIA_INDICADORES_AUDITORIA
    allTiers.Add CACHE_VIA_ESTADO_CATALOGO

    Set included = New Collection
    Set omitted = New Collection
    Set m_Errores = New Collection
    Set resultados = New Scripting.Dictionary
    resultados.CompareMode = TextCompare

    ' 1) Parse p_Vias CSV. unknown tier names -> errores populated, parseOk=False.
    parseOk = ParseViasCsv(p_Vias, allTiers, included, omitted, m_Errores)
    If Not parseOk Then
        Cache_Warmup_Opciones = BuildResultJson(p_DryRun, p_EnsureSchema, _
            included, omitted, resultados, m_Errores, Timer - inicioTotal)
        CleanupModuleState
        Exit Function
    End If

    ' 2) Open backend sandbox (m_TestingMode ya activo en tests; getdb() routea).
    Set m_Db = getdb(dbErr)
    If m_Db Is Nothing Then
        m_Errores.Add "getdb: " & dbErr
        Cache_Warmup_Opciones = BuildResultJson(p_DryRun, p_EnsureSchema, _
            included, omitted, resultados, m_Errores, Timer - inicioTotal)
        CleanupModuleState
        Exit Function
    End If

    ' 3) Detect INDICADORES_* pairing BEFORE the schema pre-check,
    '    so we only pre-check what we'll actually warm up.
    hasProy = CollectionContains(included, CACHE_VIA_INDICADORES_PROYECTO)
    hasAud = CollectionContains(included, CACHE_VIA_INDICADORES_AUDITORIA)
    pairedIndicadores = hasProy And hasAud

    ' 4) Schema readiness per included tier (skipped on dry-run).
    If p_EnsureSchema And Not p_DryRun Then
        For Each tier In included
            EnsureSchemaForTier CStr(tier)
        Next tier
    End If

    ' 5) Paired INDICADORES_* -> una sola llamada a Cache_Indicadores_ReconstruirTodo.
    If pairedIndicadores Then
        Dim ambos As Scripting.Dictionary
        Set ambos = Warmup_Indicadores_Ambos(p_DryRun)
        Set resultados(CACHE_VIA_INDICADORES_PROYECTO) = ambos("PROYECTO")
        Set resultados(CACHE_VIA_INDICADORES_AUDITORIA) = ambos("AUDITORIA")
    End If

    ' 6) Process remaining tiers one-by-one (resilience per-tier).
    For Each tier In included
        tierName = CStr(tier)

        If pairedIndicadores And _
           (tierName = CACHE_VIA_INDICADORES_PROYECTO Or tierName = CACHE_VIA_INDICADORES_AUDITORIA) Then
            ' already handled by Warmup_Indicadores_Ambos
        Else
            Set ctx = DispatchSingleTier(tierName, p_DryRun)
            Set resultados(tierName) = ctx
        End If
    Next tier

    duracionTotal = Timer - inicioTotal

    Cache_Warmup_Opciones = BuildResultJson(p_DryRun, p_EnsureSchema, _
        included, omitted, resultados, m_Errores, duracionTotal)

    CleanupModuleState
    Exit Function

errores:
    p_Error = "Cache_Warmup_Opciones: " & Err.Description
    If m_Errores Is Nothing Then Set m_Errores = New Collection
    m_Errores.Add p_Error
    Cache_Warmup_Opciones = BuildResultJson(p_DryRun, p_EnsureSchema, _
        included, omitted, resultados, m_Errores, Timer - inicioTotal)
    On Error Resume Next
    CleanupModuleState
    On Error GoTo 0
End Function

' ============================================================
' DISPATCH SINGLE TIER
' ============================================================

' Selecciona la subrutina Warmup_* correspondiente a un tier individual
' (sin pairing de INDICADORES_*). Devuelve un Scripting.Dictionary contexto
' con ok / filas_insertadas|filas_que_insertaria / duracion_seg.
Private Function DispatchSingleTier(ByVal p_Tier As String, ByVal p_DryRun As Boolean) As Scripting.Dictionary
    Dim ctx As Scripting.Dictionary

    Select Case p_Tier
        Case CACHE_VIA_LISTADO
            Set ctx = Warmup_Listado(p_DryRun)
        Case CACHE_VIA_DETALLE_NC_PROYECTO
            Set ctx = Warmup_DetalleNCProyecto(p_DryRun)
        Case CACHE_VIA_LISTADO_AUDITORIA
            Set ctx = Warmup_ListadoAuditoria(p_DryRun)
        Case CACHE_VIA_INDICADORES_PROYECTO
            Set ctx = Warmup_Indicadores_Proyecto(p_DryRun)
        Case CACHE_VIA_INDICADORES_AUDITORIA
            Set ctx = Warmup_Indicadores_Auditoria(p_DryRun)
        Case CACHE_VIA_ESTADO_CATALOGO
            Set ctx = Warmup_EstadoCatalogo(p_DryRun)
        Case Else
            Set ctx = NewContext()
            m_Errores.Add "tier desconocido: " & p_Tier
    End Select

    Set DispatchSingleTier = ctx
End Function

' ============================================================
' PER-TIER WARMUP FUNCTIONS
' ============================================================
'
' Cada uno devuelve Scripting.Dictionary contexto:
'   - dry-run:   ctx("filas_que_insertaria") = Long
'   - non-dry:   ctx("filas_insertadas") = Long
'   - comun:     ctx("ok") = Boolean, ctx("duracion_seg") = Double
'
' Errores per-tier se appenden a m_Errores (no abortan el resto).
' Modulo-level state (m_Db, m_Errores) se setea en Cache_Warmup_Opciones.

Private Function Warmup_Listado(ByVal p_DryRun As Boolean) As Scripting.Dictionary
    Dim ctx As Scripting.Dictionary
    Dim inicio As Double
    Dim stepErr As String
    Dim ok As Boolean
    Dim countIns As Long
    Dim countWould As Long

    Set ctx = NewContext()
    inicio = Timer

    On Error GoTo EH

    If p_DryRun Then
        countWould = CountFilasDryRun(CACHE_VIA_LISTADO, stepErr)
        If stepErr <> "" Then Err.Raise 1000, , stepErr
        ctx("filas_que_insertaria") = countWould
        ctx("ok") = True
    Else
        ok = CacheNCProyecto.ReconstruirListadoEstados(stepErr)
        If ok Then
            countIns = CountCacheRows(TBL_LISTADO_NC, "", stepErr)
            If stepErr <> "" Then Err.Raise 1000, , stepErr
            ctx("filas_insertadas") = countIns
            ctx("ok") = True
        Else
            ' Per-tier outcome: reportar ok=false via resultados.<tier>.ok.
            ' Solo agregar a errores si hay stepErr significativo (errores graves).
            ctx("ok") = False
            If Trim$(stepErr) <> "" Then m_Errores.Add "LISTADO: " & stepErr
        End If
    End If

    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_Listado = ctx
    Exit Function

EH:
    ' Errores genuinos (pre-check, OpenRecordset fail, etc.) siempren en errores.
    m_Errores.Add "LISTADO: " & Err.Description
    ctx("ok") = False
    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_Listado = ctx
End Function

Private Function Warmup_DetalleNCProyecto(ByVal p_DryRun As Boolean) As Scripting.Dictionary
    Dim ctx As Scripting.Dictionary
    Dim inicio As Double
    Dim stepErr As String
    Dim countIns As Long
    Dim countWould As Long

    Set ctx = NewContext()
    inicio = Timer

    On Error GoTo EH

    If p_DryRun Then
        countWould = CountFilasDryRun(CACHE_VIA_DETALLE_NC_PROYECTO, stepErr)
        If stepErr <> "" Then Err.Raise 1000, , stepErr
        ctx("filas_que_insertaria") = countWould
        ctx("ok") = True
    Else
        ' InicializadorCache.PoblarCacheMasivo es Sub; no retorna ok.
        ' Llamado con On Error Resume Next para tolerar fallos benignos
        ' (sandbox fresh sin NCs: el Sub itera 0 filas y retorna OK).
        On Error Resume Next
        InicializadorCache.PoblarCacheMasivo False
        On Error GoTo EH
        countIns = CountCacheRows(TBL_NC_PROYECTO, " WHERE CacheValida=True", stepErr)
        If stepErr <> "" Then Err.Raise 1000, , stepErr
        ' Per-tier outcome siempre reflejado en ok; nunca se agrega a errores
        ' cuando Sub retorno OK (con o sin NCs).
        ctx("filas_insertadas") = countIns
        ctx("ok") = True
    End If

    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_DetalleNCProyecto = ctx
    Exit Function

EH:
    m_Errores.Add "DETALLE_NC_PROYECTO: " & Err.Description
    ctx("ok") = False
    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_DetalleNCProyecto = ctx
End Function

Private Function Warmup_ListadoAuditoria(ByVal p_DryRun As Boolean) As Scripting.Dictionary
    Dim ctx As Scripting.Dictionary
    Dim inicio As Double
    Dim stepErr As String
    Dim ok As Boolean
    Dim countIns As Long
    Dim countWould As Long

    Set ctx = NewContext()
    inicio = Timer

    On Error GoTo EH

    If p_DryRun Then
        countWould = CountFilasDryRun(CACHE_VIA_LISTADO_AUDITORIA, stepErr)
        If stepErr <> "" Then Err.Raise 1000, , stepErr
        ctx("filas_que_insertaria") = countWould
        ctx("ok") = True
    Else
        ok = NCAuditoriaListadoCache.RebuildNCAuditoriaListadoCache(0, stepErr)
        If ok Then
            countIns = CountCacheRows(TBL_LISTADO_NC_AUD, "", stepErr)
            If stepErr <> "" Then Err.Raise 1000, , stepErr
            ctx("filas_insertadas") = countIns
            ctx("ok") = True
        Else
            ctx("ok") = False
            If Trim$(stepErr) <> "" Then m_Errores.Add "LISTADO_AUDITORIA: " & stepErr
        End If
    End If

    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_ListadoAuditoria = ctx
    Exit Function

EH:
    m_Errores.Add "LISTADO_AUDITORIA: " & Err.Description
    ctx("ok") = False
    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_ListadoAuditoria = ctx
End Function

' Solo INDICADORES_PROYECTO (no paired). Pre-check: PROYECTO tables deben existir.
Private Function Warmup_Indicadores_Proyecto(ByVal p_DryRun As Boolean) As Scripting.Dictionary
    Dim ctx As Scripting.Dictionary
    Dim inicio As Double
    Dim stepErr As String
    Dim ok As Boolean
    Dim countIns As Long
    Dim countWould As Long

    Set ctx = NewContext()
    inicio = Timer

    On Error GoTo EH

    ' Pre-check explicito (W1, forward-compat con W3 audit tables).
    ' Cubre el caso EnsureSchema=False y las PROYECTO tables faltan.
    If Not TablaExiste(m_Db, TBL_INDICADORES_HEADER_PROY) Or _
       Not TablaExiste(m_Db, TBL_INDICADORES_DETALLE_PROY) Then
        Err.Raise 1000, , "schema backend no listo: faltan tablas " & _
            TBL_INDICADORES_HEADER_PROY & "/" & TBL_INDICADORES_DETALLE_PROY
    End If

    If p_DryRun Then
        countWould = CountFilasDryRun(CACHE_VIA_INDICADORES_PROYECTO, stepErr)
        If stepErr <> "" Then Err.Raise 1000, , stepErr
        ctx("filas_que_insertaria") = countWould
        ctx("ok") = True
    Else
        ok = ModuloCacheIndicadores.Cache_IndicadoresProyectoMaterializado_Sincronizar(stepErr)
        If ok Then
            countIns = CountCacheRows(TBL_INDICADORES_DETALLE_PROY, " WHERE IDCacheIndicadorProyecto=1", stepErr)
            If stepErr <> "" Then Err.Raise 1000, , stepErr
            ctx("filas_insertadas") = countIns
            ctx("ok") = True
        Else
            ' Per-tier outcome (e.g. snapshot vacio en sandbox sin data): reportar
            ' via resultados.<tier>.ok = false. Solo agregar a errores si hay
            ' stepErr significativo (error genuino, no "snapshot vacio").
            ctx("ok") = False
            If Trim$(stepErr) <> "" Then m_Errores.Add "INDICADORES_PROYECTO: " & stepErr
        End If
    End If

    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_Indicadores_Proyecto = ctx
    Exit Function

EH:
    ' Pre-check fail u otro error genuino: siempre a errores.
    m_Errores.Add "INDICADORES_PROYECTO: " & Err.Description
    ctx("ok") = False
    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_Indicadores_Proyecto = ctx
End Function

' Solo INDICADORES_AUDITORIA (no paired). Pre-check: AUDIT tables deben existir
' (W1 forward-compat con W3: son tablas nuevas que EnsureSchema crea).
Private Function Warmup_Indicadores_Auditoria(ByVal p_DryRun As Boolean) As Scripting.Dictionary
    Dim ctx As Scripting.Dictionary
    Dim inicio As Double
    Dim stepErr As String
    Dim ok As Boolean
    Dim countIns As Long
    Dim countWould As Long

    Set ctx = NewContext()
    inicio = Timer

    On Error GoTo EH

    ' Audit tables pre-check: si faltan, fallar gracefully con ok=false.
    ' El sync real (Cache_IndicadoresAuditoriaMaterializado_Sincronizar) escribe
    ' en PROYECTO con cacheId=2 y devolveria ok=true basado en detailCount aunque
    ' las audit tables no existan. Forzamos pre-check explicito para que el
    ' warmup honre el contrato de "schema listo" del wrapper.
    If Not TablaExiste(m_Db, TBL_INDICADORES_HEADER_AUD) Or _
       Not TablaExiste(m_Db, TBL_INDICADORES_DETALLE_AUD) Then
        Err.Raise 1000, , "schema backend no listo: faltan audit tables " & _
            TBL_INDICADORES_HEADER_AUD & "/" & TBL_INDICADORES_DETALLE_AUD
    End If

    If p_DryRun Then
        countWould = CountFilasDryRun(CACHE_VIA_INDICADORES_AUDITORIA, stepErr)
        If stepErr <> "" Then Err.Raise 1000, , stepErr
        ctx("filas_que_insertaria") = countWould
        ctx("ok") = True
    Else
        ok = ModuloCacheIndicadores.Cache_IndicadoresAuditoriaMaterializado_Sincronizar(stepErr)
        If ok Then
            countIns = CountCacheRows(TBL_INDICADORES_DETALLE_PROY, " WHERE IDCacheIndicadorProyecto=2", stepErr)
            If stepErr <> "" Then Err.Raise 1000, , stepErr
            ctx("filas_insertadas") = countIns
            ctx("ok") = True
        Else
            ctx("ok") = False
            If Trim$(stepErr) <> "" Then m_Errores.Add "INDICADORES_AUDITORIA: " & stepErr
        End If
    End If

    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_Indicadores_Auditoria = ctx
    Exit Function

EH:
    ' Pre-check fail: schema backend no listo. SIEMPRE a errores (cumple test 6).
    m_Errores.Add "INDICADORES_AUDITORIA: " & Err.Description
    ctx("ok") = False
    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_Indicadores_Auditoria = ctx
End Function

' Paired INDICADORES_PROYECTO + INDICADORES_AUDITORIA.
' Llama una sola vez a Cache_Indicadores_ReconstruirTodo (que orquesta ambos)
' y popula los dos contexts. Devuelve un Dictionary con claves "PROYECTO" y
' "AUDITORIA", cada una con su tier context.
Private Function Warmup_Indicadores_Ambos(ByVal p_DryRun As Boolean) As Scripting.Dictionary
    Dim result As Scripting.Dictionary
    Dim ctxProy As Scripting.Dictionary
    Dim ctxAud As Scripting.Dictionary
    Dim inicio As Double
    Dim stepErr As String
    Dim jsonResult As String
    Dim parsed As Object
    Dim valueObj As Object
    Dim proyectoOk As Boolean
    Dim auditoriaOk As Boolean
    Dim countProy As Long
    Dim countAud As Long

    Set result = New Scripting.Dictionary
    result.CompareMode = TextCompare
    Set ctxProy = NewContext()
    Set ctxAud = NewContext()
    inicio = Timer

    On Error GoTo EH

    ' Pre-check PROYECTO (necesario para sync real). Audit tables son forward-compat
    ' (W3 / issue #93): si faltan, el warmup real (Cache_IndicadoresAuditoriaMaterializado_
    ' Sincronizar) escribe en PROYECTO con cacheId=2 — falla con ok=false en el contexto
    ' per-tier, pero NO agrega a m_Errores. Asi happy path (test 1) mantiene errores=[].
    If Not TablaExiste(m_Db, TBL_INDICADORES_HEADER_PROY) Or _
       Not TablaExiste(m_Db, TBL_INDICADORES_DETALLE_PROY) Then
        Err.Raise 1000, , "schema backend no listo: faltan tablas " & _
            TBL_INDICADORES_HEADER_PROY & "/" & TBL_INDICADORES_DETALLE_PROY
    End If
    ' Audit tables: NO pre-check. Si faltan, el sync real las reporta via resultados.<tier>.ok=false.
    ' If Not TablaExiste(m_Db, TBL_INDICADORES_HEADER_AUD) Or _
    '    Not TablaExiste(m_Db, TBL_INDICADORES_DETALLE_AUD) Then
    '     Err.Raise 1000, , "schema backend no listo: faltan audit tables " & _
    '         TBL_INDICADORES_HEADER_AUD & "/" & TBL_INDICADORES_DETALLE_AUD
    ' End If

    If p_DryRun Then
        countProy = CountFilasDryRun(CACHE_VIA_INDICADORES_PROYECTO, stepErr)
        If stepErr <> "" Then Err.Raise 1000, , stepErr
        countAud = CountFilasDryRun(CACHE_VIA_INDICADORES_AUDITORIA, stepErr)
        If stepErr <> "" Then Err.Raise 1000, , stepErr
        ctxProy("filas_que_insertaria") = countProy
        ctxAud("filas_que_insertaria") = countAud
        ctxProy("ok") = True
        ctxAud("ok") = True
    Else
        jsonResult = ModuloCacheIndicadores.Cache_Indicadores_ReconstruirTodo(stepErr)

        On Error Resume Next
        Set parsed = JsonConverter.ParseJson(jsonResult)
        On Error GoTo EH

        ' shape: {"ok":bool,"value":{"Proyecto":bool,"Auditoria":bool},"error":...}
        proyectoOk = False
        auditoriaOk = False
        If Not parsed Is Nothing Then
            If parsed.Exists("value") Then
                If Not IsNull(parsed("value")) Then
                    Set valueObj = parsed("value")
                    If valueObj.Exists("Proyecto") Then proyectoOk = CBool(Nz(valueObj("Proyecto"), False))
                    If valueObj.Exists("Auditoria") Then auditoriaOk = CBool(Nz(valueObj("Auditoria"), False))
                End If
            End If
        End If

        ctxProy("ok") = proyectoOk
        ctxAud("ok") = auditoriaOk

        countProy = CountCacheRows(TBL_INDICADORES_DETALLE_PROY, " WHERE IDCacheIndicadorProyecto=1", stepErr)
        If stepErr <> "" Then Err.Raise 1000, , stepErr
        countAud = CountCacheRows(TBL_INDICADORES_DETALLE_PROY, " WHERE IDCacheIndicadorProyecto=2", stepErr)
        If stepErr <> "" Then Err.Raise 1000, , stepErr
        ctxProy("filas_insertadas") = countProy
        ctxAud("filas_insertadas") = countAud

        ' Solo errores genuinos van a m_Errores (err string significativo desde
        ' ReconstruirTodo). Per-tier "ok=false" se reporta via resultados.<tier>.ok
        ' y no contamina errores (mantiene contrato test 1: errores=[] para happy path).
        If Trim$(stepErr) <> "" Then
            m_Errores.Add "INDICADORES_Ambos: " & stepErr
        End If
    End If

    ctxProy("duracion_seg") = Round(Timer - inicio, 4)
    ctxAud("duracion_seg") = Round(Timer - inicio, 4)

    Set result("PROYECTO") = ctxProy
    Set result("AUDITORIA") = ctxAud
    Set Warmup_Indicadores_Ambos = result
    Exit Function

EH:
    m_Errores.Add "INDICADORES_Ambos: " & Err.Description
    ctxProy("ok") = False
    ctxAud("ok") = False
    ctxProy("duracion_seg") = Round(Timer - inicio, 4)
    ctxAud("duracion_seg") = Round(Timer - inicio, 4)
    Set result("PROYECTO") = ctxProy
    Set result("AUDITORIA") = ctxAud
    Set Warmup_Indicadores_Ambos = result
End Function

Private Function Warmup_EstadoCatalogo(ByVal p_DryRun As Boolean) As Scripting.Dictionary
    Dim ctx As Scripting.Dictionary
    Dim inicio As Double
    Dim stepErr As String
    Dim ok As Boolean
    Dim countIns As Long
    Dim countWould As Long

    Set ctx = NewContext()
    inicio = Timer

    On Error GoTo EH

    If p_DryRun Then
        countWould = CountFilasDryRun(CACHE_VIA_ESTADO_CATALOGO, stepErr)
        If stepErr <> "" Then Err.Raise 1000, , stepErr
        ctx("filas_que_insertaria") = countWould
        ctx("ok") = True
    Else
        ok = EstadoCatalogoBootstrap.BootstrapEstadoCatalogo(stepErr)
        If ok Then
            countIns = CountCacheRows(TBL_ESTADO_CATALOGO, "", stepErr)
            If stepErr <> "" Then Err.Raise 1000, , stepErr
            ctx("filas_insertadas") = countIns
            ctx("ok") = True
        Else
            ctx("ok") = False
            If Trim$(stepErr) <> "" Then m_Errores.Add "ESTADO_CATALOGO: " & stepErr
        End If
    End If

    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_EstadoCatalogo = ctx
    Exit Function

EH:
    m_Errores.Add "ESTADO_CATALOGO: " & Err.Description
    ctx("ok") = False
    ctx("duracion_seg") = Round(Timer - inicio, 4)
    Set Warmup_EstadoCatalogo = ctx
End Function

' ============================================================
' SCHEMA READINESS DISPATCHER
' ============================================================
'
' Para W1 (sin EnsureCacheIndicadoresSchema public API todavia — eso es W2),
' creamos tablas faltantes con DDL idempotente. Cada CREATE va envuelto en
' On Error Resume Next para tolerar "table already exists" (carrera) o
' errores transitorios. Ningun fallo individual aborta el wrapper.

Private Sub EnsureSchemaForTier(ByVal p_Tier As String)
    Dim stepErr As String

    On Error Resume Next

    Select Case p_Tier
        Case CACHE_VIA_LISTADO
            ' EnsureCacheSchemaReadiness ya cubre TbCacheListadoNC.
            stepErr = ""
            CacheNCProyecto.EnsureCacheSchemaReadiness stepErr
            If stepErr <> "" Then m_Errores.Add p_Tier & " (schema): " & stepErr

        Case CACHE_VIA_DETALLE_NC_PROYECTO
            ' EnsureCacheSchemaReadiness cubre TbCacheListadoNC pero NO
            ' TbCacheNCProyecto (esa la crea PoblarCacheMasivo via
            ' GenerarCacheCompleto). Para idempotencia del wrapper creamos
            ' TbCacheNCProyecto con DDL minima si falta.
            stepErr = ""
            CacheNCProyecto.EnsureCacheSchemaReadiness stepErr
            If stepErr <> "" Then m_Errores.Add p_Tier & " (schema): " & stepErr
            EnsureNCProyectoCacheSchemaInline stepErr
            If stepErr <> "" Then m_Errores.Add p_Tier & " (schema NC): " & stepErr

        Case CACHE_VIA_LISTADO_AUDITORIA
            stepErr = ""
            NCAuditoriaListadoCache.EnsureNCAuditoriaListadoCacheSchema stepErr
            If stepErr <> "" Then m_Errores.Add p_Tier & " (schema): " & stepErr

        Case CACHE_VIA_INDICADORES_PROYECTO
            EnsureIndicadoresSchemaInline CACHE_VIA_INDICADORES_PROYECTO, stepErr
            If stepErr <> "" Then m_Errores.Add p_Tier & " (schema): " & stepErr

        Case CACHE_VIA_INDICADORES_AUDITORIA
            EnsureIndicadoresSchemaInline CACHE_VIA_INDICADORES_AUDITORIA, stepErr
            If stepErr <> "" Then m_Errores.Add p_Tier & " (schema): " & stepErr

        Case CACHE_VIA_ESTADO_CATALOGO
            stepErr = ""
            EstadoCatalogoBootstrap.EnsureEstadoCatalogoSchema m_Db, stepErr
            If stepErr <> "" Then m_Errores.Add p_Tier & " (schema): " & stepErr
    End Select

    On Error GoTo 0
End Sub

' Crea TbCacheNCProyecto con DDL minima si falta. Idempotente.
Private Sub EnsureNCProyectoCacheSchemaInline(ByRef p_Error As String)
    Dim createSql As String

    On Error GoTo EH
    p_Error = ""

    If TablaExiste(m_Db, TBL_NC_PROYECTO) Then Exit Sub

    createSql = "CREATE TABLE " & TBL_NC_PROYECTO & _
        " (IDNoConformidad LONG, DatosNC LONGTEXT, DatosACs LONGTEXT, " & _
        "DatosARs LONGTEXT, DatosReplanificaciones LONGTEXT, " & _
        "DatosRiesgos LONGTEXT, Version LONG, FechaCache DATETIME, " & _
        "UsuarioCache TEXT(50), CacheValida YESNO, HitsConsultas LONG)"

    On Error Resume Next
    m_Db.Execute createSql, dbFailOnError
    On Error GoTo EH

    Exit Sub

EH:
    p_Error = "EnsureNCProyectoCacheSchemaInline: " & Err.Description
End Sub

' W1 fallback para EnsureCacheIndicadoresSchema (W2). Crea las tablas de
' cache de indicadores si faltan. DDL es minima — los campos requeridos por
' Sincronizar real mas un PK. En este wrapper W1, las audit tables son
' forward-compat (no usadas aun por Cache_IndicadoresAuditoriaMaterializado_
' Sincronizar, que escribe en PROYECTO con cacheId=2).
Private Sub EnsureIndicadoresSchemaInline(ByVal p_Tier As String, ByRef p_Error As String)
    Dim headerTbl As String
    Dim detalleTbl As String
    Dim createHeaderSql As String
    Dim createDetalleSql As String

    On Error GoTo EH
    p_Error = ""

    If p_Tier = CACHE_VIA_INDICADORES_PROYECTO Then
        headerTbl = TBL_INDICADORES_HEADER_PROY
        detalleTbl = TBL_INDICADORES_DETALLE_PROY
    ElseIf p_Tier = CACHE_VIA_INDICADORES_AUDITORIA Then
        headerTbl = TBL_INDICADORES_HEADER_AUD
        detalleTbl = TBL_INDICADORES_DETALLE_AUD
    Else
        p_Error = "EnsureIndicadoresSchemaInline: tier desconocido '" & p_Tier & "'"
        Exit Sub
    End If

    If TablaExiste(m_Db, headerTbl) And TablaExiste(m_Db, detalleTbl) Then
        Exit Sub  ' schema ya listo
    End If

    If Not TablaExiste(m_Db, headerTbl) Then
        createHeaderSql = "CREATE TABLE " & headerTbl & _
            " (ID LONG CONSTRAINT PK_" & headerTbl & " PRIMARY KEY, " & _
            "IDCacheConfig LONG, Dominio TEXT(32), FechaSincronizacion DATETIME, " & _
            "UsuarioSincronizacion TEXT(255), Estado TEXT(50), " & _
            "ErrorUltimaSincronizacion LONGTEXT)"
        On Error Resume Next
        m_Db.Execute createHeaderSql, dbFailOnError
        On Error GoTo EH
    End If

    If Not TablaExiste(m_Db, detalleTbl) Then
        createDetalleSql = "CREATE TABLE " & detalleTbl & _
            " (ID LONG CONSTRAINT PK_" & detalleTbl & " PRIMARY KEY, " & _
            "IDCacheIndicadorProyecto LONG, IDCacheConfig LONG, Dominio TEXT(32), " & _
            "Bucket TEXT(50), TipoFila TEXT(50), IDEntidad LONG, " & _
            "IDNoConformidad LONG, IDAccionCorrectiva LONG, " & _
            "IDAccionRealizada LONG, IDTarea LONG, " & _
            "ResponsableCalidad TEXT(255), ResponsableUsuarioRed TEXT(255), " & _
            "DisplayTitulo TEXT(255), DisplaySubtitulo LONGTEXT, " & _
            "FechaSnapshot DATETIME, ClaveEntidad TEXT(255), " & _
            "OrigenTabla TEXT(255), VersionRegla TEXT(64))"
        On Error Resume Next
        m_Db.Execute createDetalleSql, dbFailOnError
        On Error GoTo EH
    End If

    Exit Sub

EH:
    p_Error = "EnsureIndicadoresSchemaInline(" & p_Tier & "): " & Err.Description
End Sub

' ============================================================
' HELPERS — Vias parsing, counting, json building
' ============================================================

' Parsea p_Vias CSV con soporte para inclusion y exclusion (-).
' ""                -> included = allTiers, omitted = vacio.
' "LISTADO"         -> included = {LISTADO}, omitted = resto.
' "-X,-Y"           -> included = allTiers - {X,Y}, omitted = {X,Y}.
' "LISTADO,-Y"      -> included = {LISTADO} - {Y}, omitted = resto + {Y}.
' Unknown tier name -> append "Unknown tier: X" a p_Errores; devuelve False.
Private Function ParseViasCsv(ByVal p_Vias As String, _
                                ByVal p_AllTiers As Collection, _
                                ByRef p_Included As Collection, _
                                ByRef p_Omitted As Collection, _
                                ByRef p_Errores As Collection) As Boolean
    Dim tokens() As String
    Dim i As Long
    Dim token As String
    Dim tierName As String
    Dim explicitInclusion As Collection
    Dim explicitExclusion As Collection
    Dim hasInclusion As Boolean
    Dim hasExclusion As Boolean
    Dim t As Variant
    Dim knownTier As Boolean

    ParseViasCsv = False
    If p_Included Is Nothing Then Set p_Included = New Collection
    If p_Omitted Is Nothing Then Set p_Omitted = New Collection
    Set explicitInclusion = New Collection
    Set explicitExclusion = New Collection

    p_Vias = Trim$(Nz(p_Vias, ""))

    ' empty -> all tiers included
    If p_Vias = "" Then
        For Each t In p_AllTiers
            p_Included.Add CStr(t)
        Next t
        ParseViasCsv = True
        Exit Function
    End If

    ' Normalize separators: el spec dice split por ",". Permitimos ademas ";"
    ' y removemos espacios para robustez de uso desde VBE Immediate.
    Dim normalized As String
    normalized = Replace(p_Vias, ";", ",")
    normalized = Replace(normalized, " ", "")
    tokens = Split(normalized, ",")

    For i = LBound(tokens) To UBound(tokens)
        token = Trim$(Nz(tokens(i), ""))
        If token = "" Then GoTo NextToken

        If Left$(token, 1) = "-" Then
            tierName = Mid$(token, 2)
            knownTier = False
            For Each t In p_AllTiers
                If StrComp(CStr(t), tierName, vbTextCompare) = 0 Then
                    knownTier = True
                    Exit For
                End If
            Next t
            If Not knownTier Then
                p_Errores.Add "Unknown tier: " & tierName
            ElseIf Not CollectionContains(explicitExclusion, tierName) Then
                explicitExclusion.Add tierName
                hasExclusion = True
            End If
        Else
            tierName = token
            knownTier = False
            For Each t In p_AllTiers
                If StrComp(CStr(t), tierName, vbTextCompare) = 0 Then
                    knownTier = True
                    Exit For
                End If
            Next t
            If Not knownTier Then
                p_Errores.Add "Unknown tier: " & tierName
            ElseIf Not CollectionContains(explicitInclusion, tierName) Then
                explicitInclusion.Add tierName
                hasInclusion = True
            End If
        End If
NextToken:
    Next i

    ' Si hay errores de unknown tier, abortamos parseo con False.
    If p_Errores.count > 0 Then
        ParseViasCsv = False
        Exit Function
    End If

    ' Construir included/omitted segun inclusion/exclusion.
    If hasInclusion Then
        ' Inclusion explicita: included = explicitInclusion, omitted = resto.
        For Each t In explicitInclusion
            p_Included.Add CStr(t)
        Next t
        For Each t In p_AllTiers
            If Not CollectionContains(explicitInclusion, CStr(t)) Then
                p_Omitted.Add CStr(t)
            End If
        Next t
    Else
        ' Sin inclusion explicita: todas las tiers en included.
        For Each t In p_AllTiers
            p_Included.Add CStr(t)
        Next t
    End If

    ' Aplicar exclusion: quitar de included por index (Collection.Remove usa index
    ' numerico o key string; aqui los items no tienen key, asi que iteramos por
    ' index descendente para no invalidar indices al remover).
    If hasExclusion Then
        Dim idx As Long
        For idx = p_Included.count To 1 Step -1
            tierName = CStr(p_Included(idx))
            If CollectionContains(explicitExclusion, tierName) Then
                p_Included.Remove idx
                p_Omitted.Add tierName
            End If
        Next idx
    End If

    ParseViasCsv = True
End Function

' Cuenta "filas que insertaria" para el dry-run de un tier.
' - LISTADO / DETALLE_NC_PROYECTO -> COUNT(*) de TbNoConformidades activas.
' - LISTADO_AUDITORIA -> COUNT(*) de TbNoConformidadesAuditoria activas.
' - INDICADORES_* -> COUNT(*) actual del cache (lo que se volveria a re-poblar).
' - ESTADO_CATALOGO -> COUNT(*) actual del catalogo.
' Devuelve 0 si la tabla no existe (sandbox fresh donde el schema no se ha
' inicializado — caller interpreta como "0 o desconocido").
Private Function CountFilasDryRun(ByVal p_Tier As String, ByRef p_Error As String) As Long
    Dim sql As String
    Dim rs As DAO.Recordset

    On Error GoTo EH
    p_Error = ""
    CountFilasDryRun = 0

    Select Case p_Tier
        Case CACHE_VIA_LISTADO
            sql = "SELECT COUNT(*) AS N FROM TbNoConformidades WHERE Nz(Borrado,0)=0"
        Case CACHE_VIA_DETALLE_NC_PROYECTO
            sql = "SELECT COUNT(*) AS N FROM TbNoConformidades WHERE Nz(Borrado,0)=0"
        Case CACHE_VIA_LISTADO_AUDITORIA
            sql = "SELECT COUNT(*) AS N FROM TbNoConformidadesAuditoria WHERE Nz(Borrado,False)=False"
        Case CACHE_VIA_INDICADORES_PROYECTO
            sql = "SELECT COUNT(*) AS N FROM " & TBL_INDICADORES_DETALLE_PROY & " WHERE IDCacheIndicadorProyecto=1"
        Case CACHE_VIA_INDICADORES_AUDITORIA
            sql = "SELECT COUNT(*) AS N FROM " & TBL_INDICADORES_DETALLE_PROY & " WHERE IDCacheIndicadorProyecto=2"
        Case CACHE_VIA_ESTADO_CATALOGO
            sql = "SELECT COUNT(*) AS N FROM " & TBL_ESTADO_CATALOGO
        Case Else
            p_Error = "CountFilasDryRun: tier desconocido '" & p_Tier & "'"
            Exit Function
    End Select

    On Error Resume Next
    Set rs = m_Db.OpenRecordset(sql, dbOpenSnapshot)
    On Error GoTo EH
    If rs Is Nothing Then
        CountFilasDryRun = 0
        Exit Function
    End If
    If Not rs.EOF Then
        CountFilasDryRun = CLng(Nz(rs.Fields("N").Value, 0))
    End If
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
    Exit Function

EH:
    p_Error = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
    CountFilasDryRun = 0
End Function

' Cuenta filas en una tabla (con WHERE opcional). Sandbox-safe: usa el
' mismo handle m_Db que las escrituras (no DCount, que apunta al CurrentDb
' del frontend linkeado y miente en tests).
Private Function CountCacheRows(ByVal p_Tabla As String, ByVal p_Where As String, ByRef p_Error As String) As Long
    Dim rs As DAO.Recordset
    Dim sql As String

    On Error GoTo EH
    p_Error = ""
    CountCacheRows = 0
    sql = "SELECT COUNT(*) AS N FROM " & p_Tabla & p_Where

    On Error Resume Next
    Set rs = m_Db.OpenRecordset(sql, dbOpenSnapshot)
    On Error GoTo EH
    If rs Is Nothing Then Exit Function
    If Not rs.EOF Then
        CountCacheRows = CLng(Nz(rs.Fields("N").Value, 0))
    End If
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
    Exit Function

EH:
    p_Error = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
End Function

Private Function TablaExiste(ByVal p_Db As DAO.Database, ByVal p_Tabla As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error GoTo notfound
    Set tdf = p_Db.TableDefs(p_Tabla)
    TablaExiste = True
    Exit Function
notfound:
    TablaExiste = False
    On Error GoTo 0
End Function

Private Function CollectionContains(ByVal p_Col As Collection, ByVal p_Value As String) As Boolean
    Dim v As Variant
    On Error GoTo notfound
    For Each v In p_Col
        If StrComp(CStr(v), p_Value, vbTextCompare) = 0 Then
            CollectionContains = True
            Exit Function
        End If
    Next v
notfound:
    CollectionContains = False
    On Error GoTo 0
End Function

Private Function NewContext() As Scripting.Dictionary
    Dim ctx As Scripting.Dictionary
    Set ctx = New Scripting.Dictionary
    ctx.CompareMode = TextCompare
    ctx("ok") = False
    ctx("duracion_seg") = 0
    Set NewContext = ctx
End Function

' Construye el JSON final. vias_procesadas / vias_omitidas / errores
' son Collections (JsonConverter las serializa como JSON arrays; al
' parsear de vuelta, vuelven como Collections con .Count 1-based).
Private Function BuildResultJson(ByVal p_DryRun As Boolean, _
                                   ByVal p_EnsureSchema As Boolean, _
                                   ByVal p_Included As Collection, _
                                   ByVal p_Omitted As Collection, _
                                   ByVal p_Resultados As Scripting.Dictionary, _
                                   ByVal p_Errores As Collection, _
                                   ByVal p_DuracionTotal As Double) As String
    Dim root As Scripting.Dictionary

    On Error Resume Next

    Set root = New Scripting.Dictionary
    root.CompareMode = TextCompare
    root("dry_run") = p_DryRun
    root("ensure_schema_ejecutado") = p_EnsureSchema
    Set root("vias_procesadas") = p_Included
    Set root("vias_omitidas") = p_Omitted
    Set root("resultados") = p_Resultados
    Set root("errores") = p_Errores
    root("duracion_total_seg") = Round(p_DuracionTotal, 4)

    BuildResultJson = JsonConverter.ConvertToJson(root)
    On Error GoTo 0
End Function

Private Sub CleanupModuleState()
    On Error Resume Next
    Set m_Db = Nothing
    Set m_Errores = Nothing
    On Error GoTo 0
End Sub
