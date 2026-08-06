Attribute VB_Name = "SandboxProvisioner"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: SandboxProvisioner
' RESPONSABILIDAD: Fase 2 del provisioning de sandbox - Localizar tablas linked
'
' POLÍTICA:
'   - NO se integra aún con la batería ni con EnsureSandboxReady
'   - Es un procedimiento MANUAL para validación primera
'   - Solo modifica el sandbox copy, NO la BD real
'   - No usa CondorError (harness layer)
'   - Logging via Debug.Print
'
' HARDENING v2:
'   - Flujo de cleanup determinístico: objetos DAO primero, luego DB, luego App
'   - CloseCurrentDatabase ANTES de Quit
'   - Set Nothing en orden inverso a la creación
'   - No hay Sleep/Wait - Access debe cerrar sin bloqueos
'   - Instancia visible para debugging (UserControl=True, Visible=True)
'   - Trazas exhaustivas en cada paso de open/import/close/quit
'
' FASE 2 (aislada):
'   1. Abre el sandbox copy en Access.Application aislado (VISIBLE para debugging)
'   2. Detecta linked tables via TableDef.Connect <> ""
'   3. Determina backends externos únicos
'   4. Copia backends a sidecars locales (junto al sandbox)
'   5. Para cada linked table:
'      a. Elimina la linked table del sandbox
'      b. Importa desde el sidecar via DoCmd.TransferDatabase
'   6. Cierra el Access.Application aislado:
'      a. Refresh all TableDefs
'      b. CloseCurrentDatabase
'      c. Set db = Nothing
'      d. Set app = Nothing  (implícito en el destroy)
'      e. app.Quit (solo si aún existe)
'   7. NO elimina sidecars (quedan para inspección/safety)
'
' INVOCACIÓN MANUAL DESDE Access VBA Immediate Window:
'   ? ProvisionSandbox_Fase2_Localizar()
'   Debug.Print ProvisionSandbox_Fase2_Localizar()
' ==========================================================================

' --- Constantes de Fase 2 ---
Private Const MINIMO_SANDBOX_FOLDER As String = "test_sandbox"
Private Const MINIMO_SANDBOX_FILENAME As String = "CONDOR_Sandbox.accdb"

' --- Helper de logging para el harness ---
Private Sub Phase2_Log(ByVal msg As String)
    On Error Resume Next
    Debug.Print "[SandboxProvisioner][Fase2] " & msg
    On Error GoTo 0
End Sub

' ==========================================================================
' FUNCIÓN PRINCIPAL: ProvisionSandbox_Fase2_Localizar
' ==========================================================================
'Propsito:
'  Toma el sandbox copy creado por ProvisionSandbox_Minimo() y convierte
'  sus linked tables en tablas locales usando un Access.Application aislado.
'
'Parametros:
'  Ninguno (usa rutas fijas del módulo)
'
'Retorna:
'  String con el path del sandbox localized, o "" si falló
'
'Invocación:
'  ? ProvisionSandbox_Fase2_Localizar()
' ==========================================================================
Public Function ProvisionSandbox_Fase2_Localizar() As String
    Dim strSandboxPath As String
    Dim strSidecarDir As String
    Dim appIsolated As Access.Application
    Dim dbIsolated As DAO.Database
    Dim colLinkedTables As Collection
    Dim colSidecarFiles As Collection
    Dim strResultado As String
    Dim blnTodoOk As Boolean
    
    On Error GoTo Errores_Fase2
    
    blnTodoOk = False
    Set colLinkedTables = New Collection
    Set colSidecarFiles = New Collection
    
    Phase2_Log "===== INICIO FASE 2: LOCALIZAR LINKED TABLES ====="
    
    ' --- Calcular rutas ---
    strSandboxPath = GetSandboxPath_Fase2()
    strSidecarDir = GetSandboxSidecarDir_Fase2()
    
    Phase2_Log "Sandbox path: " & strSandboxPath
    Phase2_Log "Sidecar dir: " & strSidecarDir
    
    If strSandboxPath = "" Then
        Err.Raise 20260, "SandboxProvisioner.Fase2", "No se pudo resolver la ruta del sandbox. Ejecute ProvisionSandbox_Minimo primero."
    End If
    
    ' --- Crear Access.Application aislado ---
    ' HARDENING v2: Visible=True para debugging - permite ver qué pasa dentro del Access
    ' UserControl=True para que Access se comporte como interactivo (no headless)
    Phase2_Log "Creando Access.Application aislado (visible=True para debug)..."
    Set appIsolated = New Access.Application
    appIsolated.Visible = True
    appIsolated.UserControl = True
    
    Phase2_Log ">> Access.Application creado correctamente"
    
    ' --- Abrir el sandbox copy en el Access.Application aislado ---
    Phase2_Log "Abriendo sandbox copy en instancia aislada..."
    Phase2_Log ">> Path: " & strSandboxPath
    appIsolated.OpenCurrentDatabase strSandboxPath, False
    Phase2_Log ">> Sandbox abierto OK"
    
    ' --- Obtener el DAO.Database desde el Access.Application ---
    Set dbIsolated = appIsolated.CurrentDb
    
    ' --- Detectar linked tables ---
    Phase2_Log "Detectando linked tables..."
    Set colLinkedTables = DetectLinkedTables_Fase2(dbIsolated)
    Phase2_Log "Linked tables detectadas: " & colLinkedTables.count
    
    If colLinkedTables.count = 0 Then
        Phase2_Log "ADVERTENCIA: No se detectaron linked tables. El sandbox podría no tener links."
    End If
    
    ' --- Copiar backends externos a sidecars ---
    Phase2_Log "Copiando backends externos a sidecars..."
    Call CopyExternalBackendsToSidecars_Fase2(colLinkedTables, strSidecarDir, colSidecarFiles)
    
    ' --- Reemplazar linked tables por imports locales ---
    Phase2_Log "Reemplazando linked tables por tablas locales..."
    Call ReplaceLinkedTablesWithLocalImports_Fase2(appIsolated, dbIsolated, colLinkedTables, strSidecarDir)
    
    ' --- Cerrar el Access.Application aislado ---
    ' HARDENING v2: Cleanup determinístico en orden inverso a la creación
    ' 1. Forzar refresh de TableDefs para synquear metadata
    ' 2. Liberar el DAO.Database primero (objetos hijos)
    ' 3. Cerrar el CurrentDatabase (libera el locking del archivo)
    ' 4. Quit (cierra el proceso Access)
    ' 5. Set Nothing (cleanup final del objeto)
    
    Phase2_Log "=== INICIANDO CLEANUP DETERMINÍSTICO ==="
    
    ' Paso 1: Refresh TableDefs
    Phase2_Log "[Cleanup-1] Refrescando TableDefs..."
    dbIsolated.TableDefs.Refresh
    
    ' Paso 2: Liberar DAO.Database (objeto hijo)
    Phase2_Log "[Cleanup-2] Liberando dbIsolated (DAO.Database)..."
    Set dbIsolated = Nothing
    Phase2_Log "[Cleanup-2] dbIsolated liberado"
    
    ' Paso 3: Cerrar CurrentDatabase (libera el lock del archivo)
    Phase2_Log "[Cleanup-3] Cerrando CurrentDatabase..."
    appIsolated.CloseCurrentDatabase
    Phase2_Log "[Cleanup-3] CurrentDatabase cerrada"
    
    ' Paso 4: Quit (cierra el proceso Access)
    Phase2_Log "[Cleanup-4] Ejecutando Quit..."
    appIsolated.Quit
    Phase2_Log "[Cleanup-4] Quit ejecutado"
    
    ' Paso 5: Set Nothing (implícito pero lo hacemos explícito para control)
    Phase2_Log "[Cleanup-5] Liberando appIsolated..."
    Set appIsolated = Nothing
    Phase2_Log "[Cleanup-5] appIsolated liberado"
    
    Phase2_Log "=== CLEANUP FINALIZADO ==="
    
    blnTodoOk = True
    
    Phase2_Log "===== FASE 2 FINALIZADA ====="
    Phase2_Log "Sidecars dejados en: " & strSidecarDir & " (no eliminados por seguridad)"
    Phase2_Log "Sandbox localized: " & strSandboxPath
    
    strResultado = strSandboxPath
    
    GoTo Salida_Fase2
    
Errores_Fase2:
    Phase2_Log "ERROR " & Err.Number & " - " & Err.description
    Phase2_Log ">>> ERROR DETECTADO - Ejecutando cleanup de emergencia en orden inverso"
    
    ' HARDENING v2: Cleanup de emergencia en orden determinístico
    ' SIEMPRE liberar objetos en orden inverso a su creación
    On Error Resume Next
    
    ' 1. Primero el objeto DAO (hijo)
    Phase2_Log "[Err-Cleanup] Liberando dbIsolated si existe..."
    If Not dbIsolated Is Nothing Then
        Set dbIsolated = Nothing
        Phase2_Log "[Err-Cleanup] dbIsolated liberado"
    End If
    
    ' 2. Luego el Access.Application
    Phase2_Log "[Err-Cleanup] Cerrando appIsolated si existe..."
    If Not appIsolated Is Nothing Then
        On Error Resume Next ' por si ya fue cerrada parcialmente
        appIsolated.CloseCurrentDatabase
        Phase2_Log "[Err-Cleanup] CurrentDatabase cerrada (o ya cerrada)"
        appIsolated.Quit
        Phase2_Log "[Err-Cleanup] Quit ejecutado"
        Set appIsolated = Nothing
        Phase2_Log "[Err-Cleanup] appIsolated liberado"
    End If
    
    On Error GoTo 0
    
    strResultado = ""

Salida_Fase2:
    Set colLinkedTables = Nothing
    Set colSidecarFiles = Nothing
    
    ProvisionSandbox_Fase2_Localizar = strResultado
End Function

' ==========================================================================
' HELPERS PRIVADOS DE FASE 2
' ==========================================================================

' ---
' GetSandboxPath_Fase2: Resuelve el path del sandbox de Fase 2
'   Usa la constante de TestSandbox.MINIMO_SANDBOX_FOLDER y MINIMO_SANDBOX_FILENAME
' ---
Private Function GetSandboxPath_Fase2() As String
    Dim strResult As String
    Dim fso As New FileSystemObject
    Dim strProjectPath As String
    Dim sep As String
    
    sep = "\"
    strProjectPath = CurrentProject.path
    If Right(strProjectPath, 1) <> sep Then strProjectPath = strProjectPath & sep
    
    strResult = strProjectPath & MINIMO_SANDBOX_FOLDER & sep & MINIMO_SANDBOX_FILENAME
    
    ' Verificar que existe
    If Not fso.FileExists(strResult) Then
        Phase2_Log "WARNING: Sandbox no encontrado en: " & strResult
        strResult = ""
    End If
    
    Set fso = Nothing
    GetSandboxPath_Fase2 = strResult
End Function

' ---
' GetSandboxSidecarDir_Fase2: Directorio para sidecars (mismo que sandbox)
' ---
Private Function GetSandboxSidecarDir_Fase2() As String
    Dim strProjectPath As String
    Dim sep As String
    
    sep = "\"
    strProjectPath = CurrentProject.path
    If Right(strProjectPath, 1) <> sep Then strProjectPath = strProjectPath & sep
    
    GetSandboxSidecarDir_Fase2 = strProjectPath & MINIMO_SANDBOX_FOLDER
End Function

' ---
' DetectLinkedTables_Fase2: Detecta linked tables en un DAO.Database
'   Retorna: Collection de Scripting.Dictionary {tableName, sourceTable, backendPath}
' ---
Private Function DetectLinkedTables_Fase2(db As DAO.Database) As Collection
    Dim colResult As New Collection
    Dim tdf As DAO.TableDef
    Dim intTablesTotal As Long
    Dim intTablesLinked As Long
    Dim objInfo As Scripting.Dictionary ' Declarar AL INICIO del procedimiento para evitar errores de hoisting
    
    intTablesTotal = 0
    intTablesLinked = 0
    
    For Each tdf In db.TableDefs
        ' Omitir tablas de sistema
        If Left(tdf.name, 4) = "MSys" Or Left(tdf.name, 4) = "USys" Then
            GoTo NextTdf
        End If
        
        intTablesTotal = intTablesTotal + 1
        
        ' Evaluar si es linked: Connect <> "" indica tabla linked
        Dim strConnect As String
        Dim strSourceTableName As String
        Dim blnEsLinked As Boolean
        
        strConnect = Trim(Nz_Fase2(tdf.Connect, ""))
        strSourceTableName = Trim(Nz_Fase2(tdf.sourceTableName, ""))
        
        blnEsLinked = (strConnect <> "")
        
        Phase2_Log "  Tabla: " & tdf.name & " | Connect='" & strConnect & "' | SourceTable='" & strSourceTableName & "' | Linked=" & blnEsLinked
        
        If blnEsLinked Then
            intTablesLinked = intTablesLinked + 1
            
            ' FIX v2: Crear NUEVA INSTANCIA del Dictionary en CADA iteracion
            ' y liberar la referencia anterior antes de crear la nueva.
            ' VBA "Dim X As New Y" a nivel de procedimiento solo instancia UNA vez;
            ' el objeto se reutiliza en cada iteración causando que todos los
            ' items de colResult apunten al mismo diccionario con valores de la
            ' última tabla procesada.
            If Not objInfo Is Nothing Then
                Set objInfo = Nothing
            End If
            Set objInfo = New Scripting.Dictionary
            objInfo("tableName") = tdf.name
            objInfo("sourceTable") = strSourceTableName
            objInfo("backendPath") = ExtractBackendPathFromConnect_Fase2(strConnect)
            objInfo("connectString") = strConnect ' Guardar connect completo para extracción de password
            
            colResult.Add objInfo
            Phase2_Log "    ^^ LINKED -> " & strSourceTableName & " @ " & objInfo("backendPath")
        End If
        
NextTdf:
    Next tdf
    
    Phase2_Log "RESUMEN: " & intTablesLinked & " linked tables de " & intTablesTotal & " totales"
    
    Set DetectLinkedTables_Fase2 = colResult
End Function

' ---
' ExtractBackendPathFromConnect_Fase2: Extrae el path del backend del Connect string
'   Connect string típico: ;DATABASE=\\server\share\path\file.accdb
' ---
Private Function ExtractBackendPathFromConnect_Fase2(strConnect As String) As String
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
    
    ExtractBackendPathFromConnect_Fase2 = strResult
End Function

' ---
' CopyExternalBackendsToSidecars_Fase2: Copia backends externos a sidecars locales
'   Usa DBEngine.CompactDatabase (DAO oficial) para crear sidecars SIN PASSWORD
'   así Access no pide contraseña durante la importacion.
'
'   colLinkedTables: Collection con {tableName, sourceTable, backendPath, connectString}
'   strSidecarDir: Directorio donde dejar los sidecars
'   colSidecarFiles: Collection de salida con las rutas de sidecar creadas
' ---
Private Sub CopyExternalBackendsToSidecars_Fase2( _
    colLinkedTables As Collection, _
    strSidecarDir As String, _
    colSidecarFiles As Collection)
    
    Dim fso As New FileSystemObject
    Dim colProcessedBackends As New Collection
    Dim varItem As Variant
    Dim strBackendPath As String
    Dim strSidecarPath As String
    Dim strBackendKey As String
    Dim strConnect As String
    Dim strPassword As String
    Dim blnYaProcesado As Boolean
    Dim strCheck As Variant
    Dim blnTienePassword As Boolean
    
    ' Crear directorio de sidecars si no existe
    If Not fso.FolderExists(strSidecarDir) Then
        fso.CreateFolder strSidecarDir
        Phase2_Log "Directorio de sidecars creado: " & strSidecarDir
    End If
    
    For Each varItem In colLinkedTables
        strBackendPath = varItem("backendPath")
        strConnect = varItem("connectString") ' Connect string con info de password
        
        If strBackendPath = "" Then GoTo NextLinked
        
        ' ¿Ya procesamos este backend?
        blnYaProcesado = False
        For Each strCheck In colProcessedBackends
            If CStr(strCheck) = strBackendPath Then
                blnYaProcesado = True
                Exit For
            End If
        Next strCheck
        
        If blnYaProcesado Then GoTo NextLinked
        
        ' Generar nombre sidecar único
        strBackendKey = GetBackendKeyFromPath_Fase2(strBackendPath)
        strSidecarPath = strSidecarDir & "\" & strBackendKey & "_sidecar.accdb"
        
        ' Extraer password del connect string (si existe)
        strPassword = ExtractPasswordFromConnect_Fase2(strConnect)
        blnTienePassword = (strPassword <> "")
        
        Phase2_Log "Procesando backend: " & strBackendPath
        Phase2_Log "  Password detectada en connect: " & blnTienePassword
        
        If fso.FileExists(strBackendPath) Then
            If blnTienePassword Then
                ' BACKEND CON PASSWORD: usar DBEngine.CompactDatabase
                ' para crear un sidecar SIN PASSWORD (destPassword="")
                ' Esto evita que Access pida password durante import
                Phase2_Log "  >> Usando DBEngine.CompactDatabase (metodo DAO oficial)"
                Phase2_Log "  >> Source password: [presente]"
                Phase2_Log "  >> Dest password: [vacio]"
                
                ' FIX v3: Verificar retorno; si no se crea el sidecar, ABORTAR FASE
                Dim blnSidecarOk As Boolean
                blnSidecarOk = CreateSidecarViaCompactDatabase(strBackendPath, strSidecarPath, strPassword)
                
                If Not blnSidecarOk Then
                    Phase2_Log "  >> ERROR CRITICO: No se pudo crear sidecar para backend: " & strBackendPath
                    Phase2_Log "  >> FASE ABORTADA: no se continuara con la importacion"
                    Err.Raise 20261, "SandboxProvisioner.Fase2", _
                        "FASE ABORTADA: No se pudo crear sidecar para backend: " & strBackendPath & ". Verifique que el archivo no esté en uso y que la password sea correcta."
                End If
                
                colSidecarFiles.Add strSidecarPath
                colProcessedBackends.Add strBackendPath
                Phase2_Log "  >> Sidecar (sin password) creado y verificado: " & strSidecarPath
            Else
                ' BACKEND SIN PASSWORD: FileCopy simple
                ' Verificar que no este en uso antes de copiar
                Phase2_Log "  >> Backend sin password - usando FileCopy directo"
                
                Dim blnArchivoDisponible As Boolean
                blnArchivoDisponible = True
                
                ' Intentar abrir en modo exclusivo para verificar disponibilidad
                On Error Resume Next
                Dim testFile As Integer
                testFile = FreeFile
                Open strBackendPath For Binary Access Read Lock Write As #testFile
                If Err.Number <> 0 Then
                    blnArchivoDisponible = False
                    Phase2_Log "  >> ADVERTENCIA: Archivo en uso, FileCopy podria fallar"
                    Err.Clear
                Else
                    Close #testFile
                End If
                On Error GoTo 0
                
                If blnArchivoDisponible Then
                    fso.CopyFile strBackendPath, strSidecarPath, True
                    colSidecarFiles.Add strSidecarPath
                    colProcessedBackends.Add strBackendPath
                    Phase2_Log "  >> Sidecar creado via FileCopy: " & strSidecarPath
                Else
                    ' Fallback: intentar de todas formas
                    On Error Resume Next
                    fso.CopyFile strBackendPath, strSidecarPath, True
                    If Err.Number = 0 Then
                        colSidecarFiles.Add strSidecarPath
                        colProcessedBackends.Add strBackendPath
                        Phase2_Log "  >> Sidecar creado (con archivo en uso): " & strSidecarPath
                    Else
                        Phase2_Log "  >> ERROR: No se pudo copiar archivo en uso: " & Err.description
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
        Else
            Phase2_Log "WARNING: Backend no encontrado: " & strBackendPath
        End If
        
NextLinked:
    Next varItem
    
    Set fso = Nothing
    Set colProcessedBackends = Nothing
End Sub

' ---
' CreateSidecarViaCompactDatabase: Crea sidecar sin password usando DAO CompactDatabase
'   sourcePath: Path al backend con password
'   destPath: Path al sidecar a crear (sin password)
'   sourcePassword: Password del backend (para abrirlo)
'
'   FIX v3: NO abre el source DB antes de CompactDatabase.
'   DBEngine.CompactDatabase requiere que el source esté CERRADO.
'   Se pasa la password directamente en el argumento Password de CompactDatabase.
'   Tras el intento, se verifica físicamente con FileExists(destPath).
'   Si el sidecar no existe, se lanza error y se aborna la fase.
' ---
Private Function CreateSidecarViaCompactDatabase( _
    ByVal sourcePath As String, _
    ByVal destPath As String, _
    ByVal sourcePassword As String) As Boolean
    
    Dim fso As New FileSystemObject
    Dim blnSidecarCreado As Boolean
    
    blnSidecarCreado = False
    CreateSidecarViaCompactDatabase = False
    
    On Error GoTo CompactError
    
    ' Eliminar destino anterior si existe (CompactDatabase no sobreescribe)
    If fso.FileExists(destPath) Then
        Phase2_Log "    [CompactDB] Eliminando archivo destino anterior..."
        Kill destPath
    End If
    
    Phase2_Log "    [CompactDB] Source (CLOSED): " & sourcePath
    Phase2_Log "    [CompactDB] Dest: " & destPath
    Phase2_Log "    [CompactDB] SourcePwd: [presente]"
    Phase2_Log "    [CompactDB] DestPwd: [vacio]"
    
    ' DBEngine.CompactDatabase(SrcName, DstName, DstLocale, Options, Password)
    ' Para .accdb con password, se pasa ";pwd=..." en el 5º argumento
    ' NO requiere abrir el source previamente; DAO lo abre internamente con esa password
    DBEngine.CompactDatabase sourcePath, destPath, dbLangGeneral, dbVersion120, ";pwd=" & sourcePassword
    
    Phase2_Log "    [CompactDB] CompactDatabase completado"
    
    GoTo Verificacion
    
CompactError:
    Phase2_Log "    [CompactDB] ERROR " & Err.Number & ": " & Err.description
    
    ' Fallback: si el backend esta en uso, copio sin compactar
    If Err.Number = 3031 Or Err.Number = 70 Then ' Error de password o archivo en uso
        Phase2_Log "    [CompactDB] Fallback: copiando archivo directamente (sin compactacion)"
        On Error Resume Next
        fso.CopyFile sourcePath, destPath, True
        If Err.Number = 0 Then
            Phase2_Log "    [CompactDB] Copia directa exitosa (verificar si pide password en import)"
        Else
            Phase2_Log "    [CompactDB] Copia directa fallida: " & Err.description
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    GoTo Verificacion
    
Verificacion:
    ' VERIFICACIÓN FÍSICA OBLIGATORIA: el sidecar debe existir en disco
    If fso.FileExists(destPath) Then
        blnSidecarCreado = True
        Phase2_Log "    [CompactDB] VERIFICADO: sidecar existe fisicamente: " & destPath
    Else
        blnSidecarCreado = False
        Phase2_Log "    [CompactDB] ERROR: sidecar NO existe fisicamente tras compactacion: " & destPath
    End If
    
    Set fso = Nothing
    
    ' Si no se pudo crear el sidecar, informar para que la fase se aborte
    If Not blnSidecarCreado Then
        Phase2_Log "    [CompactDB] ABORTO: No se pudo crear sidecar en: " & destPath
        CreateSidecarViaCompactDatabase = False
        Exit Function
    End If
    
    CreateSidecarViaCompactDatabase = True
End Function

' ---
' ExtractPasswordFromConnect_Fase2: Extrae el password del connect string
'   Connect stringtipico: ;DATABASE=\\server\share\path\file.accdb;PWD=miPassword
'   Retorna: password extraida o "" si no hay password
' ---
Private Function ExtractPasswordFromConnect_Fase2(strConnect As String) As String
    Dim strResult As String
    Dim posPwd As Long
    Dim posSemicolon As Long
    
    strResult = ""
    
    ' Buscar PWD= en el connect string (case insensitive)
    posPwd = InStr(UCase(strConnect), "PWD=")
    
    If posPwd > 0 Then
        ' Extraer password despues de PWD=
        strResult = Mid(strConnect, posPwd + 4)
        
        ' Buscar el proximo semicolon que delimita la password
        posSemicolon = InStr(strResult, ";")
        If posSemicolon > 0 Then
            strResult = Left(strResult, posSemicolon - 1)
        End If
        
        ' Limpiar espacios y quotes
        strResult = Trim(strResult)
        strResult = Replace(strResult, """", "")
    End If
    
    ExtractPasswordFromConnect_Fase2 = strResult
End Function

' ---
' GetBackendKeyFromPath_Fase2: Genera una key única desde el path del backend
' ---
Private Function GetBackendKeyFromPath_Fase2(strPath As String) As String
    Dim strResult As String
    Dim pos As Long
    
    ' Tomar solo el nombre del archivo sin extensión
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
    
    ' Limpiar caracteres problemáticos
    strResult = Replace(strResult, " ", "_")
    strResult = Replace(strResult, "-", "_")
    
    GetBackendKeyFromPath_Fase2 = strResult
End Function

' ---
' ReplaceLinkedTablesWithLocalImports_Fase2:
'   Elimina linked tables e importa desde sidecars usando DoCmd.TransferDatabase
'   El DoCmd corre dentro del Access.Application aislado (app)
' ---
Private Sub ReplaceLinkedTablesWithLocalImports_Fase2( _
    app As Access.Application, _
    db As DAO.Database, _
    colLinkedTables As Collection, _
    strSidecarDir As String)
    
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
        strBackendKey = GetBackendKeyFromPath_Fase2(strBackendPath)
        strSidecarPath = strSidecarDir & "\" & strBackendKey & "_sidecar.accdb"
        
        Phase2_Log "Procesando tabla: " & strTableName & " (origen: " & strSourceTable & ")"
        
        ' --- Eliminar la linked table del sandbox ---
        ' Primero: eliminar el TableDef
        On Error Resume Next
        db.TableDefs.Delete strTableName
        If Err.Number <> 0 Then
            Phase2_Log "  Nota: TableDefs.Delete devolvió error " & Err.Number & " (puede ser normal si la tabla ya no existe)"
            Err.Clear
        End If
        On Error GoTo 0
        
        db.TableDefs.Refresh
        
        ' Segundo: DROP TABLE por si quedó residuo (MsAccess guarda en cache)
        On Error Resume Next
        db.Execute "DROP TABLE [" & strTableName & "]"
        If Err.Number <> 0 Then
            Phase2_Log "  Nota: DROP TABLE devolvió error " & Err.Number & " (puede ser normal)"
            Err.Clear
        End If
        On Error GoTo 0
        
        db.TableDefs.Refresh
        
        ' --- Importar la tabla desde el sidecar usando DoCmd.TransferDatabase ---
        ' DoCmd.TransferDatabase opera sobre el CurrentDatabase del Application
        ' Por eso necesitamos el Access.Application aislado - el DoCmd debe ejecutarse
        ' dentro de ese contexto para que la tabla importada quede en el sandbox
        
        Phase2_Log "  Importando desde sidecar via DoCmd.TransferDatabase..."
        Phase2_Log "  Sidecar: " & strSidecarPath
        Phase2_Log "  Tabla origen: " & strSourceTable
        Phase2_Log "  Tabla destino: " & strTableName
        
        Phase2_Log "  >> Ejecutando DoCmd.TransferDatabase..."
        Phase2_Log "  >> Importando desde: " & strSidecarPath
        Phase2_Log "  >> Tabla origen: " & strSourceTable & " -> Tabla destino: " & strTableName
        
        ' DoCmd.TransferDatabase(acImport, acDatabaseType, dbName, acTable, source, dest, structureOnly)
        ' acImport = 0, acTable = 2, "Microsoft Access" = tipo para ACCDB
        app.DoCmd.TransferDatabase acImport, "Microsoft Access", strSidecarPath, acTable, strSourceTable, strTableName, False
        
        Phase2_Log "  >> TransferDatabase completado sin errores"
        
        ' HARDENING v2: Refrescar TableDefs INMEDIATAMENTE después del import
        ' para synquear metadata y evitar caches inconsistentes
        db.TableDefs.Refresh
        Phase2_Log "  >> TableDefs.Refresh ejecutado"
        
        Phase2_Log "  >> Tabla [" & strTableName & "] importada OK"
    Next varItem
    
End Sub

' ---
' Nz_Fase2: Helper para manejar Nulls
' ---
Private Function Nz_Fase2(valor As Variant, Optional valorDefecto As Variant = "") As Variant
    If IsNull(valor) Or IsEmpty(valor) Then
        Nz_Fase2 = valorDefecto
    Else
        Nz_Fase2 = valor
    End If
End Function

' ==========================================================================
' FUNCIÓN DE VERIFICACIÓN POST-FASE2
' ==========================================================================
' Verifica si el sandbox tiene tablas locales (no linked)
' Uso: ? SandboxProvisioner.Sandbox_HasLinkedTables("C:\path\to\test_sandbox\CONDOR_Sandbox.accdb")
' ==========================================================================
Public Function Sandbox_HasLinkedTables(Optional strSandboxPath As String = "") As Boolean
    Dim fso As New FileSystemObject
    Dim strPath As String
    Dim app As Access.Application
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim blnResult As Boolean
    
    blnResult = False
    
    ' Resolver path
    If strSandboxPath = "" Then
        strPath = GetSandboxPath_Fase2()
    Else
        strPath = strSandboxPath
    End If
    
    If strPath = "" Or Not fso.FileExists(strPath) Then
        Sandbox_HasLinkedTables = False
        Exit Function
    End If
    
    ' Abrir en Access.Application aislado
    Set app = New Access.Application
    app.OpenCurrentDatabase strPath, False
    Set db = app.CurrentDb
    
    ' Verificar linked tables
    For Each tdf In db.TableDefs
        If Left(tdf.name, 4) <> "MSys" And Left(tdf.name, 4) <> "USys" Then
            If Trim(Nz_Fase2(tdf.Connect, "")) <> "" Then
                blnResult = True
                Exit For
            End If
        End If
    Next tdf
    
    ' Cleanup
    app.CloseCurrentDatabase
    app.Quit
    Set app = Nothing
    Set db = Nothing
    Set fso = Nothing
    
    Sandbox_HasLinkedTables = blnResult
End Function

' ---
' Sandbox_TableCount: Cuenta tablas totales en el sandbox
' ---
Public Function Sandbox_TableCount(Optional strSandboxPath As String = "") As Long
    Dim fso As New FileSystemObject
    Dim strPath As String
    Dim app As Access.Application
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim lngCount As Long
    
    lngCount = 0
    
    ' Resolver path
    If strSandboxPath = "" Then
        strPath = GetSandboxPath_Fase2()
    Else
        strPath = strSandboxPath
    End If
    
    If strPath = "" Or Not fso.FileExists(strPath) Then
        Sandbox_TableCount = 0
        Exit Function
    End If
    
    ' Abrir en Access.Application aislado
    Set app = New Access.Application
    app.OpenCurrentDatabase strPath, False
    Set db = app.CurrentDb
    
    ' Contar tablas (excluyendo sistema)
    For Each tdf In db.TableDefs
        If Left(tdf.name, 4) <> "MSys" And Left(tdf.name, 4) <> "USys" Then
            lngCount = lngCount + 1
        End If
    Next tdf
    
    ' Cleanup
    app.CloseCurrentDatabase
    app.Quit
    Set app = Nothing
    Set db = Nothing
    Set fso = Nothing
    
    Sandbox_TableCount = lngCount
End Function

' ---
' Sandbox_ListTables: Lista tablas del sandbox con su tipo (local/linke
' ---
Public Function Sandbox_ListTables(Optional strSandboxPath As String = "") As String
    Dim fso As New FileSystemObject
    Dim strPath As String
    Dim app As Access.Application
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strResult As String
    
    strResult = ""
    
    ' Resolver path
    If strSandboxPath = "" Then
        strPath = GetSandboxPath_Fase2()
    Else
        strPath = strSandboxPath
    End If
    
    If strPath = "" Or Not fso.FileExists(strPath) Then
        Sandbox_ListTables = "Sandbox no encontrado"
        Exit Function
    End If
    
    ' Abrir en Access.Application aislado
    Set app = New Access.Application
    app.OpenCurrentDatabase strPath, False
    Set db = app.CurrentDb
    
    ' Listar tablas
    For Each tdf In db.TableDefs
        If Left(tdf.name, 4) <> "MSys" And Left(tdf.name, 4) <> "USys" Then
            Dim strTipo As String
            If Trim(Nz_Fase2(tdf.Connect, "")) <> "" Then
                strTipo = "LINKED"
            Else
                strTipo = "LOCAL"
            End If
            strResult = strResult & tdf.name & " (" & strTipo & ")" & vbCrLf
        End If
    Next tdf
    
    ' Cleanup
    app.CloseCurrentDatabase
    app.Quit
    Set app = Nothing
    Set db = Nothing
    Set fso = Nothing
    
    Sandbox_ListTables = strResult
End Function

' ==========================================================================
' HELPERS PÚBLICOS DE LIMPIEZA DE SIDECARS
' ==========================================================================

' ---
' CleanupSidecars_Fase2: Elimina manualmente los sidecars del sandbox
'   Uso: Call SandboxProvisioner.CleanupSidecars_Fase2()
'        ? SandboxProvisioner.CleanupSidecars_Fase2()
'
' Retorna: Cantidad de sidecars eliminados, o -1 si falló
' ---
Public Function CleanupSidecars_Fase2(Optional ByVal blnSilent As Boolean = False) As Long
    Dim fso As New FileSystemObject
    Dim strSidecarDir As String
    Dim strFile As String
    Dim colFiles As New Collection
    Dim varItem As Variant
    Dim lngDeleted As Long
    Dim lngErrors As Long
    
    lngDeleted = 0
    lngErrors = 0
    
    On Error GoTo Errores_Cleanup
    
    strSidecarDir = GetSandboxSidecarDir_Fase2()
    
    If Not fso.FolderExists(strSidecarDir) Then
        If Not blnSilent Then Phase2_Log "Directorio de sidecars no existe: " & strSidecarDir
        CleanupSidecars_Fase2 = 0
        Exit Function
    End If
    
    ' Recopilar archivos *_sidecar.accdb
    Dim fld As Folder
    Set fld = fso.GetFolder(strSidecarDir)
    
    Dim fl As File
    For Each fl In fld.Files
        If Right(LCase(fl.name), 19) = "_sidecar.accdb" Then
            colFiles.Add fl.path
        End If
    Next fl
    
    If colFiles.count = 0 Then
        If Not blnSilent Then Phase2_Log "No se encontraron sidecars para eliminar en: " & strSidecarDir
        CleanupSidecars_Fase2 = 0
        Exit Function
    End If
    
    ' Eliminar cada sidecar
    For Each varItem In colFiles
        strFile = CStr(varItem)
        On Error Resume Next
        fso.DeleteFile strFile, True
        If Err.Number = 0 Then
            lngDeleted = lngDeleted + 1
            If Not blnSilent Then Phase2_Log "Eliminado: " & strFile
        Else
            lngErrors = lngErrors + 1
            If Not blnSilent Then Phase2_Log "ERROR al eliminar " & strFile & ": " & Err.description
            Err.Clear
        End If
        On Error GoTo Errores_Cleanup
    Next varItem
    
    If Not blnSilent Then
        Phase2_Log "Cleanup completado: " & lngDeleted & " eliminados, " & lngErrors & " errores"
    End If
    
    CleanupSidecars_Fase2 = lngDeleted
    
    GoTo Salida_Cleanup
    
Errores_Cleanup:
    Phase2_Log "ERROR CleanupSidecars_Fase2 " & Err.Number & " - " & Err.description
    CleanupSidecars_Fase2 = -1
    
Salida_Cleanup:
    Set fso = Nothing
    Set colFiles = Nothing
    Set fld = Nothing
    Set fl = Nothing
End Function

' ---
' ListSidecars_Fase2: Lista los sidecars actuales del sandbox (sin eliminarlos)
'   Uso: ? SandboxProvisioner.ListSidecars_Fase2()
' Retorna: String con lista de paths de sidecar
' ---
Public Function ListSidecars_Fase2() As String
    Dim fso As New FileSystemObject
    Dim strSidecarDir As String
    Dim strResult As String
    Dim fld As Folder
    Dim fl As File
    
    strResult = ""
    strSidecarDir = GetSandboxSidecarDir_Fase2()
    
    If Not fso.FolderExists(strSidecarDir) Then
        ListSidecars_Fase2 = "(directorio de sidecars no existe)"
        Exit Function
    End If
    
    Set fld = fso.GetFolder(strSidecarDir)
    For Each fl In fld.Files
        If Right(LCase(fl.name), 19) = "_sidecar.accdb" Then
            strResult = strResult & fl.path & vbCrLf
        End If
    Next fl
    
    If strResult = "" Then
        strResult = "(sin sidecars en " & strSidecarDir & ")"
    End If
    
    Set fso = Nothing
    Set fld = Nothing
    Set fl = Nothing
    
    ListSidecars_Fase2 = strResult
End Function

