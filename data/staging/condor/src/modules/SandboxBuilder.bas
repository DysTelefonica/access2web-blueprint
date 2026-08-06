Attribute VB_Name = "SandboxBuilder"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: SandboxBuilder
' DESCRIPCIÓN: Sandbox infrastructure simplificada que reemplaza el anterior
'              enfoque complejo (SandboxGestor + Validator + Config + CloneHelper).
'
' NUEVO ENFOQUE SIMPLIFICADO:
'   1. FILECOPY del backend CONDOR a ubicación local de test
'   2. Abrir el COPY y examinar sus LINKED tables para descubrir backends externos
'   3. Copiar esos backends externos a archivos sidecar locales en el sandbox
'   4. Construir plan de importación desde las linked tables del COPY
'   5. ELIMINAR las linked tables del COPY antes de importar (evita colisiones)
'   6. Importar las tablas requeridas desde los sidecar locales via TransferDatabase
'   7. Eliminar los backends sidecar auxuliares, dejando solo el CONDOR local
'
' NOTA: Este módulo es para el ENGINE/INTERNAL. Los módulos de test harness
'       (como TestSandbox) pueden usarlo. Los tests propiamente dichos
'       NO deben usar CondorError - solo el engine sí.
' ==========================================================================

' --- Constantes ---
Private Const DB_LINKED As Long = &H80000000
Private Const CONDOR_BACKEND_NETWORK_PATH As String = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR\condor_datos.accdb"

' --- Variables privadas ---
Private m_strSandboxPath As String
Private m_strCondorBackendPath As String
Private m_strSidecarDir As String
Private m_dbSandbox As DAO.Database
Private m_colSidecarFiles As Collection

' ==========================================================================
' MÉTODOS PÚBLICOS PRINCIPALES
' ==========================================================================

' ---
' BuildSandbox: Construye o abre el sandbox de forma idempotente.
'   POLITICA PERSISTENT: Si el sandbox ya existe en disco, se ABRE y REUSA.
'   Solo hace build fresco si el sandbox no existe.
'
'   Para forzar rebuild: usar ResetSandboxIfInvalid o eliminar el archivo manualmente.
'
'   1. Resuelve paths
'   2. Si el sandbox existe: OPEN (reuse) - retorna DB abierta
'   3. Si no existe: FILECOPY + proceso de linked tables + import
'   4. Retorna el DAO.Database del sandbox
' ---
Public Function BuildSandbox() As DAO.Database
    On Error GoTo Errores
    
    Call Trace("[SandboxBuilder][BuildSandbox] INICIO")
    
    ' Inicializar colección de sidecars para cleanup
    Set m_colSidecarFiles = New Collection
    
    ' Resolver paths
    Call ResolveSandboxPaths
    Call Trace("[SandboxBuilder][BuildSandbox] Paths resueltos. SandboxPath=" & m_strSandboxPath)
    
    ' ========================================================================
    ' POLITICA PERSISTENT: Verificar si el sandbox ya existe
    ' Si existe, hacer OPEN en lugar de rebuild para evitar Error 70
    ' ========================================================================
    Dim fso As New FileSystemObject
    If fso.FileExists(m_strSandboxPath) Then
        Call Trace("[SandboxBuilder][BuildSandbox] Sandbox existe en disco. Abriendo en lugar de rebuild...")
        Set m_dbSandbox = OpenSandbox()
        Set BuildSandbox = m_dbSandbox
        Call Trace("[SandboxBuilder][BuildSandbox] FIN OK (reuse). Sandbox abierto: " & m_strSandboxPath)
        Exit Function
    End If
    
    ' ========================================================================
    ' Sandbox no existe - hacer build fresco
    ' ========================================================================
    Call Trace("[SandboxBuilder][BuildSandbox] Sandbox no existe. Haciendo build fresco...")
    
    ' Step 1: FILECOPY del backend CONDOR
    Call FileCopyCondorBackend
    Call Trace("[SandboxBuilder][BuildSandbox] Step 1: FILECOPY completado")
    
    ' Step 2: Abrir el COPY y detectar linked tables + sidecars
    Dim colLinkedTables As Collection
    Set colLinkedTables = DetectLinkedTablesAndSidecars
    Call Trace("[SandboxBuilder][BuildSandbox] Step 2: Detectadas " & colLinkedTables.count & " linked tables")
    
    ' Step 3: Copiar backends externos a sidecars
    Call CopyExternalBackendsToSidecars(colLinkedTables)
    Call Trace("[SandboxBuilder][BuildSandbox] Step 3: Sidecars creados")
    
    ' Step 4 + 5: Eliminar linked tables y importar desde sidecars
    Call ReplaceLinkedTablesWithLocalImports(colLinkedTables)
    Call Trace("[SandboxBuilder][BuildSandbox] Step 4+5: Linked tables reemplazadas por imports")
    
    ' Step 6: Eliminar sidecars auxuliares
    Call CleanupSidecars
    Call Trace("[SandboxBuilder][BuildSandbox] Step 6: Sidecars eliminados")
    
    ' El sandbox ya está abierto y listo
    Set BuildSandbox = m_dbSandbox
    
    Call Trace("[SandboxBuilder][BuildSandbox] FIN OK (build fresco). Sandbox listo en: " & m_strSandboxPath)
    Exit Function
    
Errores:
    Call Trace("[SandboxBuilder][BuildSandbox] ERROR " & Err.Number & " - " & Err.description)
    Call CleanupOnError
    Err.Raise Err.Number, "SandboxBuilder.BuildSandbox", Err.description
End Function

' ---
' OpenSandbox: Abre el sandbox existente (para reuse)
' Retorna: DAO.Database del sandbox abierto
' ---
Public Function OpenSandbox() As DAO.Database
    On Error GoTo Errores
    
    Call ResolveSandboxPaths
    
    If m_strSandboxPath = "" Then
        Err.Raise 20202, "SandboxBuilder.OpenSandbox", "Sandbox path no resuelto"
    End If
    
    Dim fso As New FileSystemObject
    If Not fso.FileExists(m_strSandboxPath) Then
        Err.Raise 20202, "SandboxBuilder.OpenSandbox", "Sandbox no existe: " & m_strSandboxPath
    End If
    
    ' Abrir con password del backend CONDOR (usa abstracción centralizada)
    ' Usar g_wsCondor si está disponible para que las transacciones del test harness funcionen
    If Not g_wsCondor Is Nothing Then
        Set m_dbSandbox = g_wsCondor.OpenDatabase(m_strSandboxPath, False, False, ";pwd=" & GetPasswordDB())
    Else
        Set m_dbSandbox = DBEngine.OpenDatabase(m_strSandboxPath, False, False, ";pwd=" & GetPasswordDB())
    End If
    
    Set OpenSandbox = m_dbSandbox
    Exit Function
    
Errores:
    Err.Raise Err.Number, "SandboxBuilder.OpenSandbox", Err.description
End Function

' ---
' CloseSandbox: Cierra el sandbox
' ---
Public Sub CloseSandbox()
    On Error Resume Next
    If Not m_dbSandbox Is Nothing Then
        m_dbSandbox.Close
        Set m_dbSandbox = Nothing
    End If
End Sub

' ---
' GetSandboxPath: Retorna la ruta del sandbox
' ---
Public Property Get GetSandboxPath() As String
    GetSandboxPath = m_strSandboxPath
End Property

' ==========================================================================
' MÉTODOS PRIVADOS - IMPLEMENTACIÓN DEL NUEVO ENFOQUE
' ==========================================================================

' ---
' ResolveSandboxPaths: Resuelve las rutas del sandbox y backends
' ---
Private Sub ResolveSandboxPaths()
    Dim fso As New FileSystemObject
    Dim strProjectPath As String
    Dim strTestSandboxDir As String
    
    ' Proyecto path
    strProjectPath = CurrentProject.path
    If Right(strProjectPath, 1) <> "\" Then strProjectPath = strProjectPath & "\"
    
    ' Directorio de sandbox
    strTestSandboxDir = strProjectPath & "test_sandbox"
    
    ' Crear directorio si no existe
    If Not fso.FolderExists(strTestSandboxDir) Then
        fso.CreateFolder strTestSandboxDir
    End If
    
    ' Path del sandbox = copy del backend CONDOR
    m_strSandboxPath = strTestSandboxDir & "\CONDOR_Test.accdb"
    m_strSidecarDir = strTestSandboxDir
    
    ' Resolver path del backend CONDOR
    ' Asumimos que el backend está junto al frontend en el mismo directorio
    Dim strFrontendPath As String
    strFrontendPath = CurrentDb.name
    
    Dim pos As Long
    pos = InStrRev(strFrontendPath, "\")
    If pos > 0 Then
        m_strCondorBackendPath = Left(strFrontendPath, pos) & "CONDOR_Backend.accdb"
    Else
        ' No se pudo resolver, fallback
        m_strCondorBackendPath = strProjectPath & "CONDOR_Backend.accdb"
    End If
    
    ' Si el backend no existe en esa ubicación, buscar pattern conocido
    If Not fso.FileExists(m_strCondorBackendPath) Then
        ' Intentar con el path del JSON config si existe
        Dim strAltPath As String
        strAltPath = strProjectPath & "test_infrastructure\config\"
        ' No modificamos m_strCondorBackendPath aquí - se intenta en Build
    End If
    
    Set fso = Nothing
End Sub

' ---
' FileCopyCondorBackend: Step 1 - FILECOPY del backend CONDOR
' ---
Private Sub FileCopyCondorBackend()
    Dim fso As New FileSystemObject
    
    ' Resolver path del backend CONDOR si no está resuelto
    If m_strCondorBackendPath = "" Or Not fso.FileExists(m_strCondorBackendPath) Then
        Call ResolveCondorBackendPath
    End If
    
    ' Verificar que existe el backend
    If Not fso.FileExists(m_strCondorBackendPath) Then
        Err.Raise 20201, "SandboxBuilder.FileCopyCondorBackend", _
            "Backend CONDOR no encontrado en: " & m_strCondorBackendPath & ". No se puede crear el sandbox."
    End If
    
    ' Eliminar sandbox anterior si existe
    If fso.FileExists(m_strSandboxPath) Then
        Call Trace("[SandboxBuilder] Eliminando sandbox previo...")
        fso.DeleteFile m_strSandboxPath, True
    End If
    
    ' FILECOPY
    Call Trace("[SandboxBuilder] FileCopy desde " & m_strCondorBackendPath & " a " & m_strSandboxPath)
    fso.CopyFile m_strCondorBackendPath, m_strSandboxPath, True
    
    Set fso = Nothing
End Sub

' ---
' ResolveCondorBackendPath: Resuelve la ruta del backend CONDOR
'   Busca en el directorio del proyecto: condor_datos.accdb (backend real)
'   NO hace fallback a CONDOR.accdb porque es el FRONTEND, no el backend.
'   Falla claramente si no encuentra el backend.
' ---
Private Sub ResolveCondorBackendPath()
    Dim fso As New FileSystemObject
    Dim strProjectPath As String
    Dim strFrontendPath As String
    Dim blnBackendEsFrontend As Boolean
    
    strProjectPath = CurrentProject.path
    If Right(strProjectPath, 1) <> "\" Then strProjectPath = strProjectPath & "\"
    
    ' Prioridad 1: backend real conocido en red
    If fso.FileExists(CONDOR_BACKEND_NETWORK_PATH) Then
        m_strCondorBackendPath = CONDOR_BACKEND_NETWORK_PATH
        Call Trace("[SandboxBuilder][ResolveCondorBackendPath] Backend resuelto (red): " & m_strCondorBackendPath)
    
    ' Prioridad 2: copia/local excepcional en el directorio del proyecto
    ElseIf fso.FileExists(strProjectPath & "condor_datos.accdb") Then
        m_strCondorBackendPath = strProjectPath & "condor_datos.accdb"
        Call Trace("[SandboxBuilder][ResolveCondorBackendPath] Backend resuelto (local fallback): " & m_strCondorBackendPath)
    Else
        ' Backend no encontrado - verificar si el fallback seria el frontend (error comun)
        strFrontendPath = strProjectPath & "CONDOR.accdb"
        blnBackendEsFrontend = fso.FileExists(strFrontendPath)
        
        Call Trace("[SandboxBuilder][ResolveCondorBackendPath] ERROR: Backend no encontrado. Se esperaba: " & CONDOR_BACKEND_NETWORK_PATH)
        Call Trace("[SandboxBuilder][ResolveCondorBackendPath] ERROR: Fallback local tampoco encontrado en: " & strProjectPath & "condor_datos.accdb")
        
        If blnBackendEsFrontend Then
            Call Trace("[SandboxBuilder][ResolveCondorBackendPath] ERROR: 'CONDOR.accdb' existe pero es el FRONTEND, no el backend. NO se usa como fallback.")
        End If
        
        ' No hacemos fallback silencioso - fallamos explicitamente
        m_strCondorBackendPath = ""  ' Marcar como no resuelto
        Err.Raise 20201, "SandboxBuilder.ResolveCondorBackendPath", _
            "No se pudo resolver el backend de CONDOR. Se esperaba: " & CONDOR_BACKEND_NETWORK_PATH
    End If
    
    Set fso = Nothing
End Sub

' ---
' DetectLinkedTablesAndSidecars: Step 2 - Abre el copy y detecta linked tables
'   También identifica qué backends externos se necesitan
'
'   METODO DE DETECCION (segun guia oficial DAO):
'     - Para linked table: TableDef.Connect contiene el string de conexion/path (NO vacio)
'     - Para base/local Access table: TableDef.Connect es zero-length string
'     - SourceTableName es informativo: local vacio, linked tiene el nombre origen
'
'   Usamos Connect <> "" (trimmed) como indicador primario, reforzado por SourceTableName.
'
' Retorna: Collection de dictionaries {tableName, sourceTable, backendPath}
' ---
Private Function DetectLinkedTablesAndSidecars() As Collection
    Dim colResult As New Collection
    Dim fso As New FileSystemObject
    
    ' Abrir el sandbox copiado
    ' Usar g_wsCondor si está disponible para que las transacciones del test harness funcionen
    If Not g_wsCondor Is Nothing Then
        Set m_dbSandbox = g_wsCondor.OpenDatabase(m_strSandboxPath, False, False, ";pwd=" & GetPasswordDB())
    Else
        Set m_dbSandbox = DBEngine.OpenDatabase(m_strSandboxPath, False, False, ";pwd=" & GetPasswordDB())
    End If
    
    ' Iterar sobre TableDefs buscando linked tables
    Dim tdf As DAO.TableDef
    Dim intTablesTotal As Long
    Dim intTablesLinked As Long
    intTablesTotal = 0
    intTablesLinked = 0
    
    For Each tdf In m_dbSandbox.TableDefs
        ' Omitir tablas de sistema
        If Left(tdf.name, 4) = "MSys" Or Left(tdf.name, 4) = "USys" Then
            GoTo NextTdf
        End If
        
        intTablesTotal = intTablesTotal + 1
        
        ' EVALUACION DE LINKED TABLE (segun guia oficial DAO):
        '   - Connect <> "" (trimmed) es el indicador primario
        '   - SourceTableName <> "" refuerza la clasificacion
        Dim strConnect As String
        Dim strSourceTableName As String
        Dim strAttribs As String
        Dim blnEsLinked As Boolean
        
        strConnect = Trim(Nz(tdf.Connect, ""))
        strSourceTableName = Trim(Nz(tdf.sourceTableName, ""))
        strAttribs = CStr(tdf.Attributes)
        
        ' Clasificacion: linked si Connect tiene contenido
        blnEsLinked = (strConnect <> "")
        
        ' TRACE detallado por tabla para debugging
        Call Trace("[SandboxBuilder][DetectLinkedTablesAndSidecars] Tabla: " & tdf.name)
        Call Trace("  Atributes= " & strAttribs & " (DB_LINKED=&H80000000)")
        Call Trace("  Connect= """ & strConnect & """ " & IIf(strConnect = "", "[VACIO - local/base]", "[NO VACIO - linked]"))
        Call Trace("  SourceTableName= """ & strSourceTableName & """" & IIf(strSourceTableName = "", "[vacío]", "[origen externo]"))
        Call Trace("  Clasificado como: " & IIf(blnEsLinked, "LINKED", "LOCAL/BASE"))
        
        If blnEsLinked Then
            intTablesLinked = intTablesLinked + 1
            
            ' Parsear el Connect string para obtener source table y backend
            Dim objInfo As New Scripting.Dictionary
            objInfo("tableName") = tdf.name
            objInfo("sourceTable") = strSourceTableName
            objInfo("backendPath") = ExtractBackendPathFromConnect(strConnect)
            
            colResult.Add objInfo
            Call Trace("[SandboxBuilder]   ^^ LINKED DETECTADA -> " & objInfo("sourceTable") & " @ " & objInfo("backendPath"))
        End If
NextTdf:
    Next tdf
    
    ' TRACE resumen de deteccion
    Call Trace("[SandboxBuilder][DetectLinkedTablesAndSidecars] RESUMEN: " & _
               intTablesLinked & " linked tables detectadas de " & intTablesTotal & " tablas totales")
    
    ' FAIL EXPLICITO si no se detectaron linked tables
    ' Esto indica que el backend copiado no tiene el contenido esperado
    If intTablesLinked = 0 Then
        Call Trace("[SandboxBuilder][DetectLinkedTablesAndSidecars] ERROR: Detectadas 0 linked tables!")
        Call Trace("[SandboxBuilder][DetectLinkedTablesAndSidecars] El backend copiado deveria contener linked tables.")
        Call Trace("[SandboxBuilder][DetectLinkedTablesAndSidecars] Sandbox abortado - no se puede continuar sin backends externos.")
        
        ' Cerrar el sandbox antes de fallar
        m_dbSandbox.Close
        Set m_dbSandbox = Nothing
        
        Err.Raise 20203, "SandboxBuilder.DetectLinkedTablesAndSidecars", _
            "Detectadas 0 linked tables en el backend copiado. " & _
            "El sandbox no puede continuar porque se esperaba encontrar tablas linked " & _
            "que referencian backends externos. Verificar que el backend CONDOR " & _
            "copiado contenga linked tables validas."
    End If
    
    Set DetectLinkedTablesAndSidecars = colResult
    Set fso = Nothing
End Function

' ---
' ExtractBackendPathFromConnect: Extrae el path del backend del Connect string
' Connect string típico: ;DATABASE=\\server\share\path\file.accdb
' ---
Private Function ExtractBackendPathFromConnect(strConnect As String) As String
    Dim strResult As String
    Dim pos As Long
    
    strResult = ""
    
    ' Buscar DATABASE= en el connect string
    pos = InStr(UCase(strConnect), "DATABASE=")
    If pos > 0 Then
        strResult = Mid(strConnect, pos + 9)
        ' Limpiar trailing semicolon o otros params
        pos = InStr(strResult, ";")
        If pos > 0 Then
            strResult = Left(strResult, pos - 1)
        End If
        ' Limpiar quotes si hay
        strResult = Replace(strResult, """", "")
    End If
    
    ExtractBackendPathFromConnect = strResult
End Function

' ---
' CopyExternalBackendsToSidecars: Step 3 - Copia backends externos a sidecars locales
' colLinkedTables: Collection de dictionaries con info de linked tables
' ---
Private Sub CopyExternalBackendsToSidecars(colLinkedTables As Collection)
    Dim fso As New FileSystemObject
    Dim colProcessedBackends As New Collection
    Dim varItem As Variant
    Dim strBackendPath As String
    Dim strSidecarPath As String
    Dim strBackendKey As String
    Dim i As Long
    
    For Each varItem In colLinkedTables
        strBackendPath = varItem("backendPath")
        If strBackendPath = "" Then GoTo NextLinked
        
        ' Ya procesamos este backend?
        Dim blnAlreadyProcessed As Boolean
        blnAlreadyProcessed = False
        Dim strCheck As Variant
        For Each strCheck In colProcessedBackends
            If CStr(strCheck) = strBackendPath Then
                blnAlreadyProcessed = True
                Exit For
            End If
        Next strCheck
        
        If blnAlreadyProcessed Then GoTo NextLinked
        
        ' Generar nombre sidecar único
        strBackendKey = GetBackendKeyFromPath(strBackendPath)
        strSidecarPath = m_strSidecarDir & "\" & strBackendKey & "_sidecar.accdb"
        
        Call Trace("[SandboxBuilder] Copiando backend externo a sidecar: " & strBackendPath & " -> " & strSidecarPath)
        
        ' FILECOPY del backend externo
        If fso.FileExists(strBackendPath) Then
            fso.CopyFile strBackendPath, strSidecarPath, True
            m_colSidecarFiles.Add strSidecarPath
            colProcessedBackends.Add strBackendPath
        Else
            Call Trace("[SandboxBuilder] WARNING: Backend externo no encontrado: " & strBackendPath)
        End If
        
NextLinked:
    Next varItem
    
    Set fso = Nothing
    Set colProcessedBackends = Nothing
End Sub

' ---
' GetBackendKeyFromPath: Genera una key única desde el path del backend
' ---
Private Function GetBackendKeyFromPath(strPath As String) As String
    Dim strResult As String
    Dim i As Long
    Dim c As String
    
    ' Tomar solo el nombre del archivo sin extensión
    Dim pos As Long
    pos = InStrRev(strPath, "\")
    If pos > 0 Then
        strResult = Mid(strPath, pos + 1)
    Else
        strResult = strPath
    End If
    
    pos = InStr(strResult, ".accdb")
    If pos > 0 Then
        strResult = Left(strResult, pos - 1)
    End If
    
    ' Limpiar caracteres problemáticos para???
    strResult = Replace(strResult, " ", "_")
    strResult = Replace(strResult, "-", "_")
    
    GetBackendKeyFromPath = strResult
End Function

' ---
' ReplaceLinkedTablesWithLocalImports: Step 4+5 - Elimina linked tables e importa desde sidecars
' ---
Private Sub ReplaceLinkedTablesWithLocalImports(colLinkedTables As Collection)
    Dim varItem As Variant
    Dim strTableName As String
    Dim strSourceTable As String
    Dim strBackendPath As String
    Dim strSidecarPath As String
    Dim strBackendKey As String
    
    For Each varItem In colLinkedTables
        strTableName = varItem("tableName")
        strSourceTable = varItem("sourceTable")
        strBackendPath = varItem("backendPath")
        
        ' Buscar el sidecar correspondiente
        strBackendKey = GetBackendKeyFromPath(strBackendPath)
        strSidecarPath = m_strSidecarDir & "\" & strBackendKey & "_sidecar.accdb"
        
        Call Trace("[SandboxBuilder] Reemplazando linked table: " & strTableName)
        
        ' Eliminar la linked table del sandbox
        On Error Resume Next
        m_dbSandbox.TableDefs.Delete strTableName
        On Error GoTo 0
        
        ' Importar la tabla desde el sidecar usando TransferDatabase
        ' TransferDatabase importará la tabla como LOCAL (no linked)
        Call Trace("[SandboxBuilder] Importando desde sidecar: " & strSidecarPath & " tabla: " & strSourceTable)
        
        VBA.DoEvents  ' Yield to let file system catch up
        
        m_dbSandbox.TableDefs.Refresh
        
        On Error Resume Next
        m_dbSandbox.Execute "DROP TABLE [" & strTableName & "]"
        On Error GoTo 0
        
        m_dbSandbox.TableDefs.Refresh
        
        ' Usar TransferDatabase para importar
        ' acImport = 0, acTable = 2
        DoCmd.TransferDatabase acImport, "Microsoft Access", strSidecarPath, acTable, strSourceTable, strTableName, False
        
        Call Trace("[SandboxBuilder] Importada tabla: " & strTableName & " desde " & strSourceTable)
        
        m_dbSandbox.TableDefs.Refresh
    Next varItem
    
End Sub

' ---
' CleanupSidecars: Step 6 - Elimina los archivos sidecar temporales
' ---
Private Sub CleanupSidecars()
    Dim fso As New FileSystemObject
    Dim strSidecarPath As String
    Dim varItem As Variant
    
    For Each varItem In m_colSidecarFiles
        strSidecarPath = CStr(varItem)
        On Error Resume Next
        If fso.FileExists(strSidecarPath) Then
            fso.DeleteFile strSidecarPath, True
            Call Trace("[SandboxBuilder] Eliminado sidecar: " & strSidecarPath)
        End If
        On Error GoTo 0
    Next varItem
    
    Set fso = Nothing
End Sub

' ---
' CleanupOnError: Cleanup en caso de error durante build
' ---
Private Sub CleanupOnError()
    On Error Resume Next
    
    ' Cerrar sandbox si está abierto
    If Not m_dbSandbox Is Nothing Then
        m_dbSandbox.Close
        Set m_dbSandbox = Nothing
    End If
    
    ' Eliminar sandbox si se creó
    Dim fso As New FileSystemObject
    If m_strSandboxPath <> "" And fso.FileExists(m_strSandboxPath) Then
        fso.DeleteFile m_strSandboxPath, True
    End If
    
    ' Eliminar sidecars
    Call CleanupSidecars
    
    Set fso = Nothing
End Sub

' ---
' Trace: Helper de logging para debugging
' ---
Private Sub Trace(ByVal msg As String)
    On Error Resume Next
    Debug.Print "[SandboxBuilder] " & msg
    Call modBattery_Canonical.Canonical_Log_Trace("[SandboxBuilder] " & msg)
    On Error GoTo 0
End Sub

' ---
' Nz: Helper para manejar Nulls
' ---
Private Function Nz(valor As Variant, Optional valorDefecto As Variant = "") As Variant
    If IsNull(valor) Or IsEmpty(valor) Then
        Nz = valorDefecto
    Else
        Nz = valor
    End If
End Function

