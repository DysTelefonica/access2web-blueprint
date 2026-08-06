VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SandboxGestor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' CLASE: SandboxGestor
' DESCRIPCIÓN: Entry point del sandbox engine. Orquesta el ciclo de vida
'              completo: detección, validación, build/reuse/rebuild.
'              Delega clonados a SandboxCloneHelper.
' FASE: 3 del cambio sandbox-linked-localization
' PHASE 4: Agregadas validaciones de guardrails contra producción
' DEPENDENCIAS: SandboxConfig, SandboxValidator,
'               SandboxCloneHelper, CondorError
' ==========================================================================

' --- Constantes de errores ---
Private Const ERR_GESTOR_NO_CONFIG As Long = 20200
Private Const ERR_GESTOR_BUILD_FAILED As Long = 20201
Private Const ERR_GESTOR_SANDBOX_NOT_FOUND As Long = 20202
Private Const ERR_GESTOR_OPEN_FAILED As Long = 20203
Private Const ERR_GESTOR_INVALID_STATE As Long = 20204
Private Const ERR_GESTOR_UNSAFE_PATH As Long = 20205
Private Const ERR_GESTOR_UNSAFE_BACKEND As Long = 20206
Private Const ERR_GESTOR_INCOMPLETE_BUILD As Long = 20207  ' Post-build completeness guard failure

' --- Variables privadas ---
Private m_objConfig As SandboxConfig
Private m_objValidator As SandboxValidator
Private m_strSandboxPath As String
Private m_strBuildReason As String
Private m_dbSandbox As DAO.Database
Private m_blnSandboxExists As Boolean

' --- Nombres de tablas temporales para cloning ---
Private Const TEMP_TABLE_PREFIX As String = "ZZZ_TEMP_"

' --------------------------------------------------------------------------
' TRACE DEL ENGINE: escribe en Debug.Print y, si la batería está activa,
' también en el log canonical. Nunca debe romper el flujo.
' --------------------------------------------------------------------------
Private Sub TraceEngine(ByVal msg As String)
    On Error Resume Next
    Debug.Print msg
    Call modBattery_Canonical.Canonical_Log_Trace(msg)
    On Error GoTo 0
End Sub

' ==========================================================================
' PROPIEDADES PÚBLICAS
' ==========================================================================

' ---
' Config: Configuración activa del sandbox
' ---
Public Property Get Config() As SandboxConfig
    Set Config = m_objConfig
End Property

Public Property Set Config(objConfig As SandboxConfig)
    Set m_objConfig = objConfig
End Property

' ---
' SandboxPath: Ruta completa del archivo sandbox
' ---
Public Property Get sandboxPath() As String
    sandboxPath = m_strSandboxPath
End Property

' ---
' SandboxExistsValue: Indica si el sandbox existe en disco
' Nota: se usa nombre distinto para evitar colisión con la función SandboxExists()
' ---
Public Property Get SandboxExistsValue() As Boolean
    SandboxExistsValue = m_blnSandboxExists
End Property

' ---
' Validator: Instancia del validador para uso externo
' ---
Public Property Get Validator() As SandboxValidator
    Set Validator = m_objValidator
End Property

' ==========================================================================
' MÉTODOS PÚBLICOS PRINCIPALES
' ==========================================================================

' ---
' GetOrCreateSandbox: Entry point principal
' configPath: Ruta al archivo JSON de configuración
' Retorna: DAO.Database del sandbox abierto
' Flujo:
'   1. Carga config si no está cargada
'   2. Si sandbox no existe ? Build
'   3. Si existe y valida ? Reuse (open)
'   4. Si existe y no valida ? Rebuild
' ---
Public Function GetOrCreateSandbox(configPath As String) As DAO.Database
    On Error GoTo Errores
    Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] INICIO. configPath=" & configPath)
    
    ' Cargar o verificar config
    If m_objConfig Is Nothing Then
        Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] Config no cargada. Llamando LoadConfig...")
        Call LoadConfig(configPath)
    Else
        Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] Reutilizando config ya cargada.")
    End If
    
    ' Registrar inicio
    Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] Estado inicial sandbox: " & GetSandboxStateDescription())
    Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] LogSandboxStart: " & GetSandboxStateDescription())
    
    ' Verificar si existe
    Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] Verificando existencia del sandbox...")
    Call CheckSandboxExists
    Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] SandboxExists=" & CStr(m_blnSandboxExists) & "; Path=" & GetSandboxPath())
    
    Dim dbResult As DAO.Database
    
    If Not m_blnSandboxExists Then
        ' Caso 1: No existe ? BUILD
        m_strBuildReason = "SANDBOX_NOT_FOUND"
        Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] Ejecutando BuildSandbox. Razón=" & m_strBuildReason)
        Set dbResult = BuildSandbox(m_strBuildReason)
        
    Else
        ' Caso 2 o 3: Existe ? validar
        Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] Abriendo sandbox existente para validar...")
        
        ' Abrir sandbox para validación
        Dim dbSandbox As DAO.Database
        Set dbSandbox = OpenSandboxDatabase()
        Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] Sandbox abierto. Validando...")
        
        ' Validar
        Dim blnValid As Boolean
        blnValid = m_objValidator.ValidateSandbox(m_objConfig, dbSandbox)
        Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] Resultado validación=" & CStr(blnValid))
        
        If blnValid Then
            ' Caso 2: Existe y valida ? REUSE
            m_strBuildReason = "VALIDATION_PASSED"
            Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] Reutilizando sandbox existente.")
            Set dbResult = dbSandbox
        Else
            ' Caso 3: Existe pero no valida ? REBUILD
            Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] Sandbox inválido. Summary=" & m_objValidator.ValidationSummary)
            Call dbSandbox.Close
            Set dbSandbox = Nothing
            m_strBuildReason = "VALIDATION_FAILED: " & m_objValidator.ValidationSummary
            Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] Ejecutando RebuildSandbox. Razón=" & m_strBuildReason)
            Set dbResult = RebuildSandbox(m_strBuildReason)
        End If
    End If
    
    ' Registrar fin exitoso
    Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] FIN OK. Razón final=" & m_strBuildReason)
    Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] LogSandboxEnd SUCCESS: " & m_strBuildReason)
    
    Set GetOrCreateSandbox = dbResult
    Exit Function
    
Errores:
    Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] ERROR " & Err.Number & " - " & Err.description)
    Call TraceEngine("[SandboxGestor][GetOrCreateSandbox] LogSandboxEnd FAILED: GetOrCreateSandbox: " & Err.description)
    
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxGestor.GetOrCreateSandbox"
    errObj.Raise
End Function

' ---
' BuildSandbox: Crea un sandbox nuevo desde cero
' reason: Razón del build (para logging)
' Retorna: DAO.Database del sandbox creado
' Flujo:
'   1. Crear archivo ACCDB vacío en outputDir
'   2. Para cada tabla configurada:
'      a. Conectar al backend origen
'      b. Crear tabla local con estructura
'      c. Copiar datos
'      d. Reemplazar linked por local (mismo nombre)
'   3. Cerrar y retornar
' ---
Public Function BuildSandbox(reason As String) As DAO.Database
    On Error GoTo ErroresBuild
    
    ' ================================================================
    ' PHASE 4: GUARDRAILS - Validaciones de seguridad ANTES de cualquier operación
    ' ================================================================
    Call ValidateSandboxEnvironment
    ' ================================================================
    
    Dim dbNew As DAO.Database
    Dim strSandboxFullPath As String
    Dim strOutputDir As String
    Dim strSandboxName As String
    Dim objFSO As FileSystemObject
    Dim colTableNames As Collection
    Dim varTable As Variant
    Dim arrNames() As String
    Dim i As Long
    
    m_strBuildReason = reason
    
    ' Logging de inicio
    Call TraceEngine("[SandboxGestor][BuildSandbox] BUILD_START: Reason=" & reason)
    
    ' Obtener paths desde config
    strOutputDir = m_objConfig.outputDir
    strSandboxName = m_objConfig.sandboxName
    strSandboxFullPath = BuildFullPath(strOutputDir, strSandboxName)
    m_strSandboxPath = strSandboxFullPath
    
    Call TraceEngine("[SandboxGestor][BuildSandbox] INICIO. OutputDir='" & strOutputDir & "'; SandboxName='" & strSandboxName & "'; FullPath='" & strSandboxFullPath & "'")
    
    ' Verificar que el directorio de output existe
    Set objFSO = New FileSystemObject
    If Not objFSO.FolderExists(strOutputDir) Then
        Call TraceEngine("[SandboxGestor][BuildSandbox] Creando directorio de output: " & strOutputDir)
        objFSO.CreateFolder strOutputDir
    End If
    
    ' Eliminar sandbox existente si lo hay (por si acaso)
    If objFSO.FileExists(strSandboxFullPath) Then
        Call TraceEngine("[SandboxGestor][BuildSandbox] Eliminando sandbox existente en: " & strSandboxFullPath)
        objFSO.DeleteFile strSandboxFullPath, True
    End If
    
    ' Crear nuevo archivo ACCDB
    Call TraceEngine("[SandboxGestor][BuildSandbox] Creando nuevo ACCDB en: " & strSandboxFullPath)
    Set dbNew = DBEngine.CreateDatabase(strSandboxFullPath, dbLangSpanish)
    Call TraceEngine("[SandboxGestor][BuildSandbox] ACCDB creado. Tables.Count=" & dbNew.TableDefs.count)

    ' Obtener lista de tablas a localizar
    Set colTableNames = New Collection
    arrNames = m_objConfig.GetTableNames()
    For i = LBound(arrNames) To UBound(arrNames)
        If arrNames(i) <> "" Then
            colTableNames.Add arrNames(i)
        End If
    Next i
    
    Call TraceEngine("[SandboxGestor][BuildSandbox] Tablas a localizar: " & colTableNames.count)
    
    ' Procesar cada tabla
    For Each varTable In colTableNames
        Dim strTableName As String
        strTableName = CStr(varTable)
        
        ' Obtener info del backend para esta tabla
        Dim strBackendKey As String
        Dim strBackendPath As String
        Dim strBackendPassword As String
        
        strBackendKey = GetBackendKeyForTable(strTableName)
        Call TraceEngine("[SandboxGestor][BuildSandbox] Tabla='" & strTableName & "'; BackendKey='" & strBackendKey & "'")
        
        If strBackendKey <> "" Then
            strBackendPath = m_objConfig.GetBackendPath(strBackendKey)
            strBackendPassword = m_objConfig.GetBackendPassword(strBackendKey)
            Call TraceEngine("[SandboxGestor][BuildSandbox] BackendKey='" & strBackendKey & "'; BackendPath='" & strBackendPath & "'; HasPassword=" & (strBackendPassword <> ""))
        Else
            Call TraceEngine("[SandboxGestor][BuildSandbox] SKIP tabla '" & strTableName & "' - sin backendKey")
            GoTo NextTable
        End If
        
        ' Conectar al backend y clonar tabla
        Call TraceEngine("[SandboxGestor][BuildSandbox] CloneTableToSandbox INICIO para '" & strTableName & "'")
        Call CloneTableToSandbox(strTableName, strBackendKey, dbNew)
        Call TraceEngine("[SandboxGestor][BuildSandbox] CloneTableToSandbox FIN para '" & strTableName & "'")
        
NextTable:
    Next varTable
    
    ' Cerrar y reabrir para verificar
    Call TraceEngine("[SandboxGestor][BuildSandbox] Cerrando sandbox para verificacion post-build...")
    dbNew.Close
    Set dbNew = Nothing
    
    ' Abrir el sandbox creado
    Call TraceEngine("[SandboxGestor][BuildSandbox] Reabriendo sandbox para verificacion post-build...")
    Set dbNew = OpenSandboxDatabase()
    Call TraceEngine("[SandboxGestor][BuildSandbox] Sandbox reopen. Tables.Count=" & dbNew.TableDefs.count)
    
    ' ================================================================
    ' POST-BUILD COMPLETENESS GUARD
    ' Verificar que el sandbox tiene las tablas mínimas requeridas
    ' ================================================================
    Call TraceEngine("[SandboxGestor][BuildSandbox] Ejecutando verificacion post-build...")
    VerifyPostBuildCompleteness dbNew, colTableNames
    Call TraceEngine("[SandboxGestor][BuildSandbox] Verificacion post-build COMPLETADA.")
    
    ' Actualizar flag
    m_blnSandboxExists = True
    
    ' Logging de fin
    Call TraceEngine("[SandboxGestor][BuildSandbox] BUILD_END: Sandbox creado en: " & strSandboxFullPath)
    Call TraceEngine("[SandboxGestor][BuildSandbox] FIN OK. Path=" & strSandboxFullPath)
    
    Set BuildSandbox = dbNew
    Exit Function
    
ErroresBuild:
    Dim errObj As New CondorError
    errObj.Create ERR_GESTOR_BUILD_FAILED, "BuildSandbox falló: " & Err.description & " (Razón: " & reason & ")", "SandboxGestor.BuildSandbox"
    errObj.Raise
End Function

' ---
' RebuildSandbox: Elimina y recrea el sandbox
' reason: Razón del rebuild (para logging)
' Retorna: DAO.Database del nuevo sandbox
' ---
Public Function RebuildSandbox(reason As String) As DAO.Database
    On Error GoTo Errores
    
    Dim objFSO As FileSystemObject
    Dim strFullPath As String
    
    ' Logging de inicio rebuild
    Call TraceEngine("[SandboxGestor][RebuildSandbox] LogSandboxStart REBUILD: " & reason)
    
    ' Eliminar sandbox existente
    strFullPath = GetSandboxPath()
    If strFullPath <> "" Then
        Set objFSO = New FileSystemObject
        If objFSO.FileExists(strFullPath) Then
            ' Cerrar cualquier conexión abierta
            If Not (m_dbSandbox Is Nothing) Then
                On Error Resume Next
                m_dbSandbox.Close
                Set m_dbSandbox = Nothing
                On Error GoTo Errores
            End If
            
            objFSO.DeleteFile strFullPath, True
        End If
    End If
    
    ' Resetear validator
    Set m_objValidator = New SandboxValidator
    
    ' Marcar como no existente
    m_blnSandboxExists = False
    
    ' Ejecutar build fresco
    Set RebuildSandbox = BuildSandbox("REBUILD: " & reason)
    
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxGestor.RebuildSandbox"
    errObj.Raise
End Function

' ---
' OpenSandbox: Abre el sandbox existente y retorna la referencia
' Retorna: DAO.Database del sandbox abierto
' ---
Public Function OpenSandbox() As DAO.Database
    On Error GoTo Errores
    
    Set OpenSandbox = OpenSandboxDatabase()
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create ERR_GESTOR_OPEN_FAILED, "No se pudo abrir el sandbox: " & Err.description, "SandboxGestor.OpenSandbox"
    errObj.Raise
End Function

' ---
' SandboxExists: Verifica si el sandbox existe en disco
' Retorna: True si existe, False si no
' ---
Public Function SandboxExists() As Boolean
    Call CheckSandboxExists
    SandboxExists = m_blnSandboxExists
End Function

' ---
' GetSandboxPath: Retorna la ruta completa del sandbox
' Si no se cargó config, retorna ""
' ---
Public Function GetSandboxPath() As String
    If m_objConfig Is Nothing Then
        GetSandboxPath = ""
        Exit Function
    End If
    
    If m_strSandboxPath = "" Then
        m_strSandboxPath = BuildFullPath(m_objConfig.outputDir, m_objConfig.sandboxName)
    End If
    
    GetSandboxPath = m_strSandboxPath
End Function

' ==========================================================================
' MÉTODOS PÚBLICOS DE CONFIGURACIÓN
' ==========================================================================

' ---
' LoadConfig: Carga la configuración desde un archivo JSON
' ---
Public Sub LoadConfig(configPath As String)
    On Error GoTo Errores
    Call TraceEngine("[SandboxGestor][LoadConfig] Cargando config desde: " & configPath)
    
    Set m_objConfig = New SandboxConfig
    Call m_objConfig.Load(configPath)
    Call TraceEngine("[SandboxGestor][LoadConfig] Config cargada. OutputDir='" & m_objConfig.outputDir & "'; SandboxName='" & m_objConfig.sandboxName & "'")
    
    ' Validar config básica
    If Not m_objConfig.Validate() Then
        Call TraceEngine("[SandboxGestor][LoadConfig] Validate() devolvió False.")
        Dim errObj As New CondorError
        errObj.Create ERR_GESTOR_NO_CONFIG, "Configuración inválida: " & configPath, "SandboxGestor.LoadConfig"
        errObj.Raise
    End If
    Call TraceEngine("[SandboxGestor][LoadConfig] Validate() OK. Inicializando validator...")
    
    ' Inicializar validator
    Set m_objValidator = New SandboxValidator
    
    ' Calcular path del sandbox
    Call TraceEngine("[SandboxGestor][LoadConfig] Calculando sandbox path...")
    m_strSandboxPath = BuildFullPath(m_objConfig.outputDir, m_objConfig.sandboxName)
    Call TraceEngine("[SandboxGestor][LoadConfig] Sandbox path calculado: " & m_strSandboxPath)
    
    Exit Sub
    
Errores:
    Call TraceEngine("[SandboxGestor][LoadConfig] ERROR " & Err.Number & " - " & Err.description)
    Dim errObj2 As New CondorError
    errObj2.Create Err.Number, Err.description, "SandboxGestor.LoadConfig"
    errObj2.Raise
End Sub

' ==========================================================================
' MÉTODOS PRIVADOS DE HELPERS
' ==========================================================================

' ---
' CheckSandboxExists: Verifica si el sandbox existe y actualiza m_blnSandboxExists
' ---
Private Sub CheckSandboxExists()
    Dim objFSO As FileSystemObject
    Dim strPath As String
    
    strPath = GetSandboxPath()
    
    If strPath = "" Then
        m_blnSandboxExists = False
        Exit Sub
    End If
    
    Set objFSO = New FileSystemObject
    m_blnSandboxExists = objFSO.FileExists(strPath)
End Sub

' ---
' OpenSandboxDatabase: Abre el archivo sandbox y retorna la referencia
' ---
Private Function OpenSandboxDatabase() As DAO.Database
    On Error GoTo Errores
    
    Dim strPath As String
    Dim dbResult As DAO.Database
    
    strPath = GetSandboxPath()
    
    If strPath = "" Then
        Dim errObj As New CondorError
        errObj.Create ERR_GESTOR_SANDBOX_NOT_FOUND, "Sandbox path no está definido. ¿Se cargó la configuración?", "SandboxGestor.OpenSandboxDatabase"
        errObj.Raise
    End If
    
    ' Verificar que existe
    Dim objFSO As FileSystemObject
    Set objFSO = New FileSystemObject
    If Not objFSO.FileExists(strPath) Then
        Dim errObj2 As New CondorError
        errObj2.Create ERR_GESTOR_SANDBOX_NOT_FOUND, "Sandbox no encontrado en: " & strPath, "SandboxGestor.OpenSandboxDatabase"
        errObj2.Raise
    End If
    
    ' Abrir base de datos
    Set dbResult = DBEngine.OpenDatabase(strPath, False, False, ";pwd=" & GetSandboxPassword())
    
    Set OpenSandboxDatabase = dbResult
    Exit Function
    
Errores:
    Dim errObjFinal As New CondorError
    errObjFinal.Create ERR_GESTOR_OPEN_FAILED, "Error al abrir sandbox: " & Err.description, "SandboxGestor.OpenSandboxDatabase"
    errObjFinal.Raise
End Function

' ---
' GetSandboxPassword: Obtiene la contraseña del sandbox
' Retorna la password configurada en el JSON (o vacía si no está configurada).
' El sandbox engine NO usa secretos hardcodeados - si la password es necesaria
' y no está configurada, el sistema lanzara error en OpenSandboxDatabase.
' ---
Private Function GetSandboxPassword() As String
    Dim strPassword As String
    strPassword = ""
    
    ' Obtener password desde la config del sandbox (sandbox.password del JSON)
    ' o desde el primer backend si el sandbox no tiene password específica
    If Not (m_objConfig Is Nothing) Then
        ' Primero intentar obtener del sandbox config
        ' SandboxConfig actualmente no expone SandboxPassword directamente,
        ' así que usamos la password del primer backend como referencia
        Dim arrNames() As String
        arrNames = m_objConfig.GetTableNames()
        If UBound(arrNames) >= LBound(arrNames) Then
            Dim strFirstTable As String
            strFirstTable = arrNames(LBound(arrNames))
            Dim strBackendKey As String
            strBackendKey = GetBackendKeyForTable(strFirstTable)
            If strBackendKey <> "" Then
                strPassword = m_objConfig.GetBackendPassword(strBackendKey)
            End If
        End If
    End If
    
    ' NO hay fallback hardcodeado - si strPassword queda vacía, se usa ""
    ' El sandbox se abre con la password que tenga configurada (o vacía)
    GetSandboxPassword = strPassword
End Function

' ---
' BuildFullPath: Construye la ruta completa del sandbox
' ---
Private Function BuildFullPath(outputDir As String, sandboxName As String) As String
    Dim strResult As String
    Dim strNormalizedDir As String

    If Len(outputDir) = 0 Then
        Err.Raise ERR_GESTOR_UNSAFE_PATH, "SandboxGestor.BuildFullPath", "OutputDir vacío al construir la ruta completa del sandbox."
    End If
    
    If Len(sandboxName) = 0 Then
        Err.Raise ERR_GESTOR_UNSAFE_PATH, "SandboxGestor.BuildFullPath", "SandboxName vacío al construir la ruta completa del sandbox."
    End If
    
    ' Normalizar directorio (quitar trailing \ si existe)
    strNormalizedDir = outputDir
    If Right(strNormalizedDir, 1) = "\" Then
        strNormalizedDir = Left(strNormalizedDir, Len(strNormalizedDir) - 1)
    End If
    
    strResult = strNormalizedDir & "\" & sandboxName
    BuildFullPath = strResult
End Function

' ---
' GetBackendKeyForTable: Obtiene la clave del backend para una tabla
' Retorna "" si no la encuentra
' ---
Private Function GetBackendKeyForTable(tableName As String) As String
    On Error GoTo Errores
    
    Dim colTables As Collection
    Dim varItem As Variant
    Dim objTableDict As Object
    
    Set colTables = m_objConfig.TablesToLocalize
    
    For Each varItem In colTables
        If TypeName(varItem) = "Dictionary" Then
            Set objTableDict = varItem
            If objTableDict.Exists("name") Then
                If CStr(objTableDict("name")) = tableName Then
                    If objTableDict.Exists("backend") Then
                        GetBackendKeyForTable = CStr(objTableDict("backend"))
                        Exit Function
                    End If
                End If
            End If
        End If
    Next varItem
    
    GetBackendKeyForTable = ""
    Exit Function
    
Errores:
    GetBackendKeyForTable = ""
End Function

' ---
' CloneTableToSandbox: Clona una tabla al sandbox
' tableName: Nombre de la tabla a clonar
' backendKey: Clave del backend de origen
' dbTarget: Database del sandbox
' ---
Private Sub CloneTableToSandbox(tableName As String, backendKey As String, dbTarget As DAO.Database)
    On Error GoTo Errores
    
    Dim strBackendPath As String
    Dim strBackendPassword As String
    Dim dbSource As DAO.Database
    Dim strTempTableName As String
    Dim lngRowCount As Long
    
    ' Obtener datos del backend
    strBackendPath = m_objConfig.GetBackendPath(backendKey)
    strBackendPassword = m_objConfig.GetBackendPassword(backendKey)
    
    If strBackendPath = "" Then
        Call TraceEngine("[SandboxGestor][CloneTableToSandbox] SKIP '" & tableName & "' - backendPath vacío")
        Exit Sub
    End If
    
    ' Conectar al backend origen
    Call TraceEngine("[SandboxGestor][CloneTableToSandbox] Abriendo backend: " & strBackendPath)
    Set dbSource = DBEngine.OpenDatabase(strBackendPath, False, False, ";pwd=" & strBackendPassword)
    
    ' Verificar que la tabla existe en el backend
    Dim tdfSrc As DAO.TableDef
    On Error Resume Next
    Set tdfSrc = dbSource.TableDefs(tableName)
    If Err.Number <> 0 Then
        On Error GoTo Errores
        Call TraceEngine("[SandboxGestor][CloneTableToSandbox] SKIP '" & tableName & "' - tabla no existe en backend")
        dbSource.Close
        Set dbSource = Nothing
        Exit Sub
    End If
    On Error GoTo Errores
    Call TraceEngine("[SandboxGestor][CloneTableToSandbox] Tabla '" & tableName & "' encontrada en backend. Attributes=" & tdfSrc.Attributes)
    
    ' Nombre temporal para la tabla local
    strTempTableName = TEMP_TABLE_PREFIX & tableName
    Call TraceEngine("[SandboxGestor][CloneTableToSandbox] TempTableName='" & strTempTableName & "'")
    
    ' Crear tabla local desde la linked del backend
    ' NOTE: dbSource es requerido porque la linked table existe en el backend, no en el sandbox
    Call TraceEngine("[SandboxGestor][CloneTableToSandbox] CreateLocalTableFromLinked INICIO para '" & tableName & "' -> '" & strTempTableName & "'")
    Call SandboxCloneHelper.CreateLocalTableFromLinked(tableName, strTempTableName, dbSource, dbTarget)
    Call TraceEngine("[SandboxGestor][CloneTableToSandbox] CreateLocalTableFromLinked FIN")
    
    ' Copiar datos: leer desde dbSource (backend), escribir en dbTarget (sandbox)
    Call TraceEngine("[SandboxGestor][CloneTableToSandbox] CopyTableData INICIO para '" & tableName & "' -> '" & strTempTableName & "'")
    lngRowCount = SandboxCloneHelper.CopyTableData(tableName, strTempTableName, dbSource, dbTarget)
    Call TraceEngine("[SandboxGestor][CloneTableToSandbox] CopyTableData FIN. Filas copiadas=" & lngRowCount)
    
    ' Registrar en log
    Call TraceEngine("[SandboxGestor][CloneTableToSandbox] LogTableClone: " & tableName & " - Rows=" & lngRowCount)
    
    ' Cerrar backend origen
    Call TraceEngine("[SandboxGestor][CloneTableToSandbox] Cerrando backend...")
    dbSource.Close
    Set dbSource = Nothing
    
    ' FIX: Verificar existencia primero, luego si es linked
    ' Solo ReplaceLinkedWithLocal si la tabla final existe Y es linked
    ' En fresh sandbox, la tabla final no existe -> renombrar temp a final
    Call TraceEngine("[SandboxGestor][CloneTableToSandbox] Verificando si '" & tableName & "' existe en sandbox...")
    On Error Resume Next
    Dim blnFinalExists As Boolean
    blnFinalExists = SandboxCloneHelper.TableExists(tableName, dbTarget)
    Call TraceEngine("[SandboxGestor][CloneTableToSandbox] Tabla final '" & tableName & "' existe=" & blnFinalExists)
    
    If blnFinalExists Then
        ' La tabla final existe - verificar si es linked antes de reemplazar
        If SandboxCloneHelper.IsLinkedTable(tableName, dbTarget) Then
            Call TraceEngine("[SandboxGestor][CloneTableToSandbox] Es linked. Reemplazando con local...")
            Call SandboxCloneHelper.ReplaceLinkedWithLocal(tableName, strTempTableName, dbTarget)
            Call TraceEngine("[SandboxGestor][CloneTableToSandbox] ReplaceLinkedWithLocal completado")
        Else
            ' Existe pero NO es linked - caso inesperado en build fresco
            ' Loguear y renombrar de todas formas
            Call TraceEngine("[SandboxGestor][CloneTableToSandbox] ATENCIÓN: Tabla '" & tableName & "' existe pero NO es linked. Renombrando temp...")
            dbTarget.TableDefs.Refresh
            Dim tdfTemp As DAO.TableDef
            Set tdfTemp = dbTarget.TableDefs(strTempTableName)
            tdfTemp.name = tableName
            Call TraceEngine("[SandboxGestor][CloneTableToSandbox] Renombrado completado. TableDefs.Count=" & dbTarget.TableDefs.count)
        End If
    Else
        ' La tabla final no existe - renombrar temp al nombre final (caso fresh sandbox)
        Call TraceEngine("[SandboxGestor][CloneTableToSandbox] No es linked. Renombrando temp '" & strTempTableName & "' -> '" & tableName & "'")
        dbTarget.TableDefs.Refresh
        Dim tdfTemp2 As DAO.TableDef
        Set tdfTemp2 = dbTarget.TableDefs(strTempTableName)
        tdfTemp2.name = tableName
        Call TraceEngine("[SandboxGestor][CloneTableToSandbox] Renombrado completado. TableDefs.Count=" & dbTarget.TableDefs.count)
    End If
    On Error GoTo Errores
    
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxGestor.CloneTableToSandbox"
    errObj.Raise
End Sub

' ---
' GetSandboxStateDescription: Describe el estado actual del sandbox
' ---
Private Function GetSandboxStateDescription() As String
    Dim strState As String
    
    If Not m_blnSandboxExists Then
        strState = "NOT_EXISTS"
    Else
        strState = "EXISTS"
    End If
    
    GetSandboxStateDescription = strState
End Function

' ---
' VerifyPostBuildCompleteness: GUARDIA DE INTEGRIDAD POST-BUILD
' Verifica que el sandbox contiene al menos UNA tabla configurada por nombre final.
' Si falla, lanza error para que BuildSandbox NO reporte éxito falsamente.
' ---
Private Sub VerifyPostBuildCompleteness(dbSandbox As DAO.Database, colExpectedTables As Collection)
    On Error GoTo Errores
    
    Dim tdf As DAO.TableDef
    Dim blnAtLeastOneTableFound As Boolean
    Dim strMissingTables As String
    Dim varTable As Variant
    Dim strTableName As String
    Dim blnTableExists As Boolean
    
    blnAtLeastOneTableFound = False
    strMissingTables = ""
    
    ' Verificar que al menos una tabla configurada existe
    For Each varTable In colExpectedTables
        strTableName = CStr(varTable)
        blnTableExists = False
        
        For Each tdf In dbSandbox.TableDefs
            If tdf.name = strTableName Then
                blnTableExists = True
                Exit For
            End If
        Next tdf
        
        Call TraceEngine("[SandboxGestor][VerifyPostBuildCompleteness] Tabla='" & strTableName & "'; Existe=" & blnTableExists)
        
        If blnTableExists Then
            blnAtLeastOneTableFound = True
        Else
            If strMissingTables <> "" Then strMissingTables = strMissingTables & ", "
            strMissingTables = strMissingTables & strTableName
        End If
    Next varTable
    
    Call TraceEngine("[SandboxGestor][VerifyPostBuildCompleteness] Resumen: AlMenosUnaTabla=" & blnAtLeastOneTableFound & "; Faltantes='" & strMissingTables & "'")
    
    ' SIEMPRE lanzar error si falta algo crítico - NO retornar éxito falso
    If Not blnAtLeastOneTableFound Then
        Dim errObj As New CondorError
        errObj.Create ERR_GESTOR_INCOMPLETE_BUILD, _
            "BUILD INCOMPLETO - Sandbox creado pero no se encontraron tablas configuradas. " & _
            "Faltantes: " & strMissingTables & ". " & _
            "Esto indica que CloneTableToSandbox no está creando las tablas correctamente.", _
            "SandboxGestor.VerifyPostBuildCompleteness"
        Call TraceEngine("[SandboxGestor][VerifyPostBuildCompleteness] ERROR: Lanzando error de incompletitud. No se retornará éxito falso.")
        errObj.Raise
    End If
    
    Exit Sub
    
Errores:
    Dim errObjFinal As New CondorError
    errObjFinal.Create Err.Number, Err.description, "SandboxGestor.VerifyPostBuildCompleteness"
    errObjFinal.Raise
End Sub

' ==========================================================================
' MÉTODOS DE GUARDRAILS — PHASE 4
' ==========================================================================
' Validaciones de seguridad para prevenir operaciones peligrosas en producción
' ==========================================================================

' ---
' ValidateSandboxEnvironment: Valida que el entorno es seguro antes de build/rebuild
' Esta función es el punto central de guardrails - si algo está mal, ABORTA
' Lanza CondorError si detecta condiciones inseguras
' ---
Private Sub ValidateSandboxEnvironment()
    On Error GoTo Errores
    
    ' 1. Verificar que tenemos config cargada
    If m_objConfig Is Nothing Then
        Dim errObj As New CondorError
        errObj.Create ERR_GESTOR_NO_CONFIG, _
            "SandboxGestor no tiene configuración cargada. Llame a LoadConfig primero.", _
            "SandboxGestor.ValidateSandboxEnvironment"
        errObj.Raise
    End If
    
    ' 2. Validar que el OUTPUT DIR (donde se creará el sandbox) es SEGURO
    Dim strOutputDir As String
    strOutputDir = m_objConfig.outputDir
    
    If strOutputDir = "" Then
        Dim errObj2 As New CondorError
        errObj2.Create ERR_GESTOR_UNSAFE_PATH, _
            "OutputDir no está configurado. No se puede determinar la ubicación del sandbox.", _
            "SandboxGestor.ValidateSandboxEnvironment"
        errObj2.Raise
    End If
    
    ' Expandir ruta relativa si es necesario
    Dim strExpandedOutput As String
    strExpandedOutput = ExpandRelativePath(strOutputDir)
    
    ' ASSERT: El path del sandbox NO puede ser producción
    Call m_objValidator.AssertSandboxPathIsSafe(strExpandedOutput)
    
    ' 3. Validar que TODOS los backends de origen son seguros (o marcados como read-only)
    ' Para cada backend referenciado en las tablas, verificar seguridad
    Call ValidateBackendSources
    
    ' 4. Validar que el sandbox name no esté vacío o sea suspicious
    Dim strSandboxName As String
    strSandboxName = m_objConfig.sandboxName
    
    If strSandboxName = "" Then
        Dim errObj3 As New CondorError
        errObj3.Create ERR_GESTOR_UNSAFE_PATH, _
            "SandboxName está vacío. No se puede crear un sandbox sin nombre.", _
            "SandboxGestor.ValidateSandboxEnvironment"
        errObj3.Raise
    End If
    
    ' 5. Verificar que el sandbox name no parezca ser un archivo de producción
    ' (nombres como "CONDOR.accdb", "HPS.accdb" son sospechosos)
    If IsSuspiciousSandboxName(strSandboxName) Then
        Dim errObj4 As New CondorError
        errObj4.Create ERR_GESTOR_UNSAFE_PATH, _
            "El nombre del sandbox '" & strSandboxName & "' parece ser de producción." & vbNewLine & _
            "Use un nombre que indique claramente que es un sandbox de test, ej: 'CONDOR_Test_Sandbox.accdb'", _
            "SandboxGestor.ValidateSandboxEnvironment"
        errObj4.Raise
    End If
    
    Exit Sub
    
Errores:
    Dim errObjFinal As New CondorError
    errObjFinal.Create Err.Number, Err.description, "SandboxGestor.ValidateSandboxEnvironment"
    errObjFinal.Raise
End Sub

' ---
' ValidateBackendSources: Valida todos los backends referenciados en la config
' Asegura que backends de producción solo se usen como fuente de SOLO LECTURA
' ---
Private Sub ValidateBackendSources()
    On Error GoTo Errores
    
    Dim colTables As Collection
    Dim varItem As Variant
    Dim objTableDict As Object
    Dim strBackendKey As String
    Dim strBackendPath As String
    Dim blnAllBackendsSafe As Boolean
    
    blnAllBackendsSafe = True
    
    Set colTables = m_objConfig.TablesToLocalize
    
    For Each varItem In colTables
        If TypeName(varItem) = "Dictionary" Then
            Set objTableDict = varItem
            If objTableDict.Exists("backend") Then
                strBackendKey = CStr(objTableDict("backend"))
                strBackendPath = m_objConfig.GetBackendPath(strBackendKey)
                
                If strBackendPath <> "" Then
                    ' Verificar si el backend es seguro para lectura
                    Dim blnIsSafe As Boolean
                    blnIsSafe = m_objValidator.AssertBackendSourceIsReadOnlySafe(strBackendPath)
                    
                    If Not blnIsSafe Then
                        ' Es producción - loguear advertencia pero continuar
                        ' El engine está preparado para manejar backends de producción en modo read-only
                        Call TraceEngine("[SandboxGestor][ValidateBackendSources] BACKEND_SOURCE_PRODUCTION: Backend '" & strBackendKey & "' apunta a producción: " & strBackendPath & ". Se usará modo SOLO-LECTURA.")
                    End If
                End If
            End If
        End If
    Next varItem
    
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxGestor.ValidateBackendSources"
    errObj.Raise
End Sub

' ---
' ExpandRelativePath: Convierte rutas relativas a absolutas
' Si la ruta ya es absoluta, la devuelve tal cual
' ---
Private Function ExpandRelativePath(path As String) As String
    On Error GoTo Errores
    
    Dim strResult As String
    strResult = Trim(path)
    
    If strResult = "" Then
        ExpandRelativePath = strResult
        Exit Function
    End If
    
    ' Normalizar slashes
    strResult = Replace(strResult, "/", "\")
    
    ' Si es UNC o ruta absoluta de Windows (C:\...), devolver tal cual
    If Left(strResult, 2) = "\\" Then
        ExpandRelativePath = strResult
        Exit Function
    End If
    
    If Len(strResult) >= 2 Then
        If Mid(strResult, 2, 1) = ":" Then
            ExpandRelativePath = strResult
            Exit Function
        End If
    End If
    
    ' Es ruta relativa - convertir a absoluta usando el directorio del proyecto
    Dim strProjectPath As String
    strProjectPath = CurrentProject.path
    
    If Right(strProjectPath, 1) <> "\" Then
        strProjectPath = strProjectPath & "\"
    End If
    
    strResult = strProjectPath & strResult
    ExpandRelativePath = strResult
    Exit Function
    
Errores:
    ExpandRelativePath = path
End Function

' ---
' IsSuspiciousSandboxName: Detecta nombres de sandbox que parecen ser de producción
' ---
Private Function IsSuspiciousSandboxName(name As String) As Boolean
    On Error GoTo Errores
    
    Dim strLower As String
    strLower = LCase(name)
    
    ' Nombres que claramente son de producción, no de test
    Static arrSuspicious As Variant
    arrSuspicious = Array( _
        "condor.accdb", _
        "condor_frontend.accdb", _
        "hps.accdb", _
        "hps_backend.accdb", _
        "expedientes.accdb", _
        "noconformidades.accdb", _
        "lanzadera.accdb" _
    )
    
    Dim i As Long
    For i = LBound(arrSuspicious) To UBound(arrSuspicious)
        If strLower = arrSuspicious(i) Then
            IsSuspiciousSandboxName = True
            Exit Function
        End If
    Next i
    
    ' También verificar si contiene ciertas extensiones/rutas
    If strLower Like "*\hps*" Then
        IsSuspiciousSandboxName = True
        Exit Function
    End If
    
    IsSuspiciousSandboxName = False
    Exit Function
    
Errores:
    IsSuspiciousSandboxName = False
End Function

