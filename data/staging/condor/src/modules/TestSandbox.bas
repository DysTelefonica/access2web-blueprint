Attribute VB_Name = "TestSandbox"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: TestSandbox.bas
' RESPONSABILIDAD: Centralizar el ciclo de vida del sandbox de pruebas.
'
' Phase 1:
'   - Ruta fija relativa al proyecto (no user-specific)
'   - Clonar base de datos
'   - Abrir instancia de prueba
'   - Cerrar instancia de prueba
'
' Phase 2:
'   - Transacciones por test con rollback seguro
' ==========================================================================

' --- CONSTANTES DEL SANDBOX ---
Private Const SANDBOX_FILENAME As String = "TEST_CONDOR_Sandbox.accdb"
Private Const SANDBOX_FOLDER As String = "test_sandbox" ' Subcarpeta relativa al proyecto

' --- ESTADO DEL SANDBOX ---
Private m_SandboxPath As String          ' Ruta completa al archivo clone
Private m_OriginalPath As String         ' Ruta de la BD original
Private m_SandboxWS As DAO.Workspace    ' Workspace aislado para el sandbox
Private m_SandboxDB As DAO.Database     ' Conexión al sandbox
Private m_TransactionActive As Boolean  ' Flag: transacción en curso
Private m_TransactionDepth As Long      ' Contador de transacciones anidadas
Private m_blnSandboxExists As Boolean   ' Flag: indica si el sandbox existe en disco

' ---
' BACKEND_SOURCE: Path de red conocido del backend CONDOR
' ---
Private Const MINIMO_BACKEND_SOURCE As String = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR\condor_datos.accdb"
Private Const MINIMO_SANDBOX_FILENAME As String = "CONDOR_Sandbox.accdb"
Private Const MINIMO_SANDBOX_FOLDER As String = "test_sandbox"

' --------------------------------------------------------------------------
' HARNESS LOG: deja traza en Debug.Print y, si la batería canonical ya abrió
' el log, también en test_results.log. Nunca debe romper el flujo.
' --------------------------------------------------------------------------
Private Sub Sandbox_LogHarness(ByVal msg As String)
    On Error Resume Next
    ' Debug.Print msg
    Call modBattery_Canonical.Canonical_Log_Trace(msg)
    On Error GoTo 0
End Sub

' ==========================================================================
' FASE 1: CICLO DE VIDA DEL SANDBOX
' ==========================================================================

' --------------------------------------------------------------------------
' SANDBOX: Ruta proyecto-relativa
'   Calcula la ruta al archivo de sandbox en una subcarpeta "test_sandbox"
'   dentro del directorio del proyecto. Si la carpeta no existe, la crea.
' --------------------------------------------------------------------------
Public Function Sandbox_Ruta() As String
    Dim proyectoPath As String
    Dim sep As String
    
    ' Determinar el separador según el SO (Access puede correr en Windows)
    sep = "\"
    
    ' Obtener el directorio del proyecto desde la BD actual
    proyectoPath = CurrentDb().name
    Dim pos As Long: pos = InStrRev(proyectoPath, sep)
    If pos > 0 Then
        proyectoPath = Left(proyectoPath, pos)
    Else
        ' Fallback: usar el directorio de la aplicación
        proyectoPath = Application.CurrentProject.path
        If Right(proyectoPath, 1) <> sep Then proyectoPath = proyectoPath & sep
    End If
    
    ' Construir ruta al directorio de sandbox
    Dim sandboxFolder As String
    sandboxFolder = proyectoPath & SANDBOX_FOLDER
    
    ' Crear la carpeta si no existe
    Dim fso As New FileSystemObject
    If Not fso.FolderExists(sandboxFolder) Then
        fso.CreateFolder sandboxFolder
    End If
    
    Sandbox_Ruta = sandboxFolder & sep & SANDBOX_FILENAME
End Function

' --------------------------------------------------------------------------
' SANDBOX: Verifica si el sandbox existe
' --------------------------------------------------------------------------
Public Function Sandbox_Existe() As Boolean
    On Error Resume Next
    Sandbox_Existe = fso.FileExists(Sandbox_Ruta())
End Function

' --------------------------------------------------------------------------
' SANDBOX: Clona la base de datos original al sandbox
'   1. Determina la ruta original desde getdb()
'   2. Borra cualquier sandbox previo
'   3. Copia la BD original al sandbox
'
' OBSOLETO: Este método pertenece al enfoque antiguo (SandboxGestor).
'   No es llamado por la batería canonical (que usa EnsureSandboxReady).
'   Se mantiene para compatibilidad externa hasta que sea seguro eliminar.
' --------------------------------------------------------------------------
Public Sub Sandbox_Clonar()
    ' Debug.Print "   [Sandbox] Clonando base de datos..."
    Dim dbOrigen As DAO.Database
    
    ' Determinar la ruta original
    Set dbOrigen = getdb()
    If dbOrigen Is Nothing Then
        Err.Raise 513, "Sandbox_Clonar", "No se pudo resolver la base de datos original para crear el sandbox."
    End If
    
    m_OriginalPath = dbOrigen.name
    
    ' Calcular ruta del sandbox
    m_SandboxPath = Sandbox_Ruta()
    
    ' Eliminar sandbox anterior si existe
    If Sandbox_Existe() Then
        fso.DeleteFile m_SandboxPath, True
    End If
    
    ' Clonar la base de datos
    fso.CopyFile m_OriginalPath, m_SandboxPath, True
    
    ' Debug.Print "   [Sandbox] Clon completada: " & m_SandboxPath
End Sub

' --------------------------------------------------------------------------
' SANDBOX: Abre la instancia de prueba
'   Usa el workspace aislado g_wsCondor si está disponible,
'   o crea uno nuevo si es necesario.
'
' OBSOLETO: No es llamado por la batería canonical (que usa EnsureSandboxReady).
'   Se mantiene para compatibilidad externa y como fallback del property
'   Sandbox_DB si alguien lo llama sin haber usado EnsureSandboxReady.
' --------------------------------------------------------------------------
Public Sub Sandbox_AbrirInstancia()
    ' Debug.Print "   [Sandbox] Abriendo instancia de prueba..."
    
    If m_SandboxPath = "" Then
        Err.Raise 513, "Sandbox_AbrirInstancia", "Debe llamar a Sandbox_Clonar primero."
    End If
    
    ' Usar el workspace aislado si existe, si no crear uno nuevo
    Dim ws As DAO.Workspace
    If Not g_wsCondor Is Nothing Then
        Set ws = g_wsCondor
    Else
        Set ws = getWorkspace()
    End If
    Set m_SandboxWS = ws
    
    ' Abrir la base de datos del sandbox
    ' Contraseña de la BD de CONDOR: dpddpd
    Set m_SandboxDB = ws.OpenDatabase(m_SandboxPath, False, False, "MS Access;PWD=" & GetPasswordDB())
    Set g_dbTestInstance = m_SandboxDB
    
    ' Debug.Print "   [Sandbox] Instancia abierta: " & m_SandboxPath
End Sub

' --------------------------------------------------------------------------
' SANDBOX: Cierra la instancia de prueba
' --------------------------------------------------------------------------
Public Sub Sandbox_CerrarInstancia()
    ' Debug.Print "   [Sandbox] Cerrando instancia de prueba..."
    
    On Error Resume Next
    
    ' Cerrar la conexión si está abierta
    If Not m_SandboxDB Is Nothing Then
        m_SandboxDB.Close
        Set m_SandboxDB = Nothing
    End If
    
    ' Limpiar la referencia global del test instance
    Set g_dbTestInstance = Nothing
    
    ' Debug.Print "   [Sandbox] Instancia cerrada."
End Sub

' --------------------------------------------------------------------------
' SANDBOX: Elimina el archivo de sandbox
' --------------------------------------------------------------------------
Public Sub Sandbox_Eliminar()
    ' Debug.Print "   [Sandbox] Eliminando sandbox..."
    
    On Error Resume Next
    
    Sandbox_CerrarInstancia
    
    ' Usar m_SandboxPath directamente (establecido por EnsureSandboxReady)
    ' No llamar a Sandbox_Existe() porque usa Sandbox_Ruta() que tiene
    ' un nombre de archivo diferente al del SandboxBuilder
    If m_SandboxPath <> "" And fso.FileExists(m_SandboxPath) Then
        fso.DeleteFile m_SandboxPath, True
    End If
    
    ' Debug.Print "   [Sandbox] Sandbox eliminado."
End Sub

' ==========================================================================
' FASE 2: TRANSACCIONES POR TEST CON ROLLBACK SEGURO
' ==========================================================================

' --------------------------------------------------------------------------
' TEST: Inicia una transacción en el workspace del sandbox
'   Usa transactions anidadas via m_TransactionDepth
' --------------------------------------------------------------------------
Public Sub Test_StartTransaction()
    If Not m_SandboxWS Is Nothing Then
        m_SandboxWS.BeginTrans
        m_TransactionDepth = m_TransactionDepth + 1
        m_TransactionActive = True
        ' Debug.Print "   [Sandbox] Transacción iniciada. Depth: " & m_TransactionDepth
    Else
        Err.Raise 513, "Test_StartTransaction", "Sandbox no está abierto."
    End If
End Sub

' --------------------------------------------------------------------------
' TEST: Hace Rollback de la transacción actual
'   Solo hace rollback si m_TransactionDepth es 1 (última transacción)
'   Si hay anidamiento, decrementa el contador
' --------------------------------------------------------------------------
Public Sub Test_RollbackTransaction()
    If Not m_SandboxWS Is Nothing And m_TransactionActive Then
        If m_TransactionDepth <= 1 Then
            m_SandboxWS.Rollback
            m_TransactionDepth = 0
            m_TransactionActive = False
            ' Debug.Print "   [Sandbox] Rollback realizado."
        Else
            m_TransactionDepth = m_TransactionDepth - 1
            ' Debug.Print "   [Sandbox] Rollback pendiente (transacciones anidadas). Depth: " & m_TransactionDepth
        End If
    End If
End Sub

' --------------------------------------------------------------------------
' TEST: Confirma la transacción actual
' --------------------------------------------------------------------------
Public Sub Test_CommitTransaction()
    If Not m_SandboxWS Is Nothing And m_TransactionActive Then
        If m_TransactionDepth <= 1 Then
            m_SandboxWS.CommitTrans dbForceOSFlush
            m_TransactionDepth = 0
            m_TransactionActive = False
            ' Debug.Print "   [Sandbox] Commit realizado."
        Else
            m_TransactionDepth = m_TransactionDepth - 1
            ' Debug.Print "   [Sandbox] Commit pendiente (transacciones anidadas). Depth: " & m_TransactionDepth
        End If
    End If
End Sub

' --------------------------------------------------------------------------
' TEST: Helper para cleanup forzado (en caso de errores en tests)
' --------------------------------------------------------------------------
Public Sub Test_ForceRollback()
    On Error Resume Next
    If Not m_SandboxWS Is Nothing And m_TransactionActive Then
        m_SandboxWS.Rollback
        m_TransactionDepth = 0
        m_TransactionActive = False
        ' Debug.Print "   [Sandbox] ForceRollback ejecutado."
    End If
End Sub

' ==========================================================================
' HELPERS DE ESTADO
' ==========================================================================

' --------------------------------------------------------------------------
' ProvisionSandbox_Manual: Provisiona el sandbox para inspección manual.
'   - No dispara la batería de tests.
'   - Usa Debug.Print para logging (no depende de Canonical_Log_Trace).
'   - Retorna el path del sandbox y el DAO.Database abierto.
'
' Uso: Call ProvisionSandbox_Manual(dbOut) para obtener el database,
'      o inspeccionar SandboxBuilder.GetSandboxPath directamente.
' --------------------------------------------------------------------------
Public Function ProvisionSandbox_Manual(Optional ByRef dbOut As DAO.Database = Nothing) As String
    Dim dbResult As DAO.Database
    Dim strPath As String
    
    On Error GoTo ErroresProvision
    
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Manual] ===== INICIO PROVISION MANUAL ====="
    
    ' Verificar que SandboxBuilder está disponible
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Manual] Verificando modulo SandboxBuilder..."
    
    ' BuildSandbox es idempotente: abre existente o construye nuevo
    ' Si el sandbox ya existe, lo abre; si no, hace build fresco
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Manual] Invocando SandboxBuilder.BuildSandbox..."
    Set dbResult = SandboxBuilder.BuildSandbox()
    
    strPath = SandboxBuilder.GetSandboxPath()
    
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Manual] Sandbox listo en: " & strPath
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Manual] Database abierto: " & IIf(dbResult Is Nothing, "NO", "SI")
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Manual] ===== PROVISION FINALIZADA ====="
    
    ' Devolver por parámetro ByRef si se solicitó
    Set dbOut = dbResult
    
    ' También actualizar estado interno del módulo por compatibilidad
    Set m_SandboxDB = dbResult
    m_SandboxPath = strPath
    m_blnSandboxExists = True
    
    ProvisionSandbox_Manual = strPath
    Exit Function
    
ErroresProvision:
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Manual] ERROR " & err.Number & " - " & err.description
    Err.Raise Err.Number, "TestSandbox.ProvisionSandbox_Manual", Err.description
End Function

Public Property Get Sandbox_TransactionActive() As Boolean
    Sandbox_TransactionActive = m_TransactionActive
End Property

Public Property Get Sandbox_DB() As DAO.Database
    Dim dummy As String
    On Error GoTo ErroresValidacion
    
    ' CASO 1: Sandbox no fue abierto nunca
    If m_SandboxDB Is Nothing Then
        If m_SandboxPath = "" Then
            Err.Raise 513, "TestSandbox.Sandbox_DB", "Sandbox no fue configurado. Llame a Sandbox_Clonar y Sandbox_AbrirInstancia primero."
        End If
        ' El sandbox fue configurado (clonado) pero no abierto - abrir ahora
        Sandbox_AbrirInstancia
    Else
        ' CASO 2: Verificar si el objeto DAO sigue siendo válido
        ' Error 3420 "objeto no válido" ocurre cuando VBA garbage collector libera el objeto
        ' aunque aún tengamos una referencia. Usar .Name fuerza una validación real.
        On Error Resume Next
        dummy = m_SandboxDB.name
        If Err.Number <> 0 Then
            ' El objeto DAO fue liberado por el garbage collector - re-abrir automáticamente
            Err.Clear
            On Error GoTo ErroresValidacion
            ' Debug.Print "   [Sandbox] WARNING: Database object was garbage collected. Re-opening..."
            
            ' Re-abrir usando la misma lógica que Sandbox_AbrirInstancia
            Call Sandbox_CerrarInstancia  ' Limpia referencias residuales
            Sandbox_AbrirInstancia
            
            ' Debug.Print "   [Sandbox] Database re-opened successfully."
        Else
            On Error GoTo ErroresValidacion
        End If
    End If
    
    Set Sandbox_DB = m_SandboxDB
    Exit Property
    
ErroresValidacion:
    Err.Raise 513, "TestSandbox.Sandbox_DB", "Sandbox database inválido: " & Err.description & ". Ejecute Canonical_Setup nuevamente."
End Property

Public Property Get Sandbox_Workspace() As DAO.Workspace
    Set Sandbox_Workspace = m_SandboxWS
End Property

' ==========================================================================
' SANDBOX ENGINE WRAPPERS (enfoque simplificado)
' ==========================================================================
' Wrappers públicos para el sandbox engine simplificado (SandboxBuilder).
' El enfoque nuevo usa FILECOPY + TransferDatabase en lugar del complejo
' ciclo build/reuse/rebuild del engine anterior.
' ==========================================================================

' ---
' HasLinkedTables: Verifica si una DB tiene tablas linked activas
' dbTarget: Database a verificar (si Nothing, usa el sandbox actual)
' Retorna: True si encuentra al menos una tabla linked (no de sistema)
' ---
Public Function HasLinkedTables(Optional dbTarget As DAO.Database = Nothing) As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fso As New FileSystemObject
    
    HasLinkedTables = False
    
    ' Usar database especificada o el sandbox actual
    If dbTarget Is Nothing Then
        Set db = m_SandboxDB
    Else
        Set db = dbTarget
    End If
    
    If db Is Nothing Then Exit Function
    
    ' Iterar sobre las TableDefs
    For Each tdf In db.TableDefs
        ' Omitir tablas de sistema
        If Left(tdf.name, 4) = "MSys" Or Left(tdf.name, 4) = "USys" Then
            GoTo NextTable
        End If
        ' Verificar si es linked: el flag dbAttached está presente
        If (tdf.Attributes And &H80000000) = &H80000000 Then
            HasLinkedTables = True
            Exit For
        End If
NextTable:
    Next tdf
    
    Set fso = Nothing
End Function

' ---
' GetSandboxConfigPath: Resuelve la ruta absoluta al JSON de configuración
' del sandbox engine, relativa al proyecto.
' Retorna: Ruta completa al archivo JSON
' ---
Public Function GetSandboxConfigPath() As String
    Dim proyectoPath As String
    Dim sep As String
    Dim configRelativa As String
    
    sep = "\"
    
    ' Obtener directorio del proyecto
    proyectoPath = CurrentProject.path
    If Right(proyectoPath, 1) <> sep Then
        proyectoPath = proyectoPath & sep
    End If
    
    ' Ruta relativa al JSON de configuración
    configRelativa = "test_infrastructure" & sep & "config" & sep & "sandbox-linked-localization.json"
    
    GetSandboxConfigPath = proyectoPath & configRelativa
End Function

' ---
' EnsureSandboxReady: read-only compatibility shim (PR2, Spec-007).
' Delegates to TestHelper.ForceLocalBackend, which validates the asserted
' sandbox WITHOUT writing to TbConfiguracionBackends. The previous
' implementation mutated BackendActivo = "TEST" and BackendTest = <path>;
' that contract is retired — production configuration rows are read-only
' for test code (Spec-007).
' Retorna: DAO.Database del sandbox abierto y listo
' Si el sandbox no puede prepararse, lanza error claro.
' ---
Public Function EnsureSandboxReady() As DAO.Database
    Dim cfgError As String

    On Error GoTo ErroresEnsure

    ' Asegurar workspace aislado disponible para transacciones del test harness
    If g_wsCondor Is Nothing Then
        Set g_wsCondor = getWorkspace()
    End If
    Set m_SandboxWS = g_wsCondor

    ' PR2: read-only routing. No more BackendActivo/BackendTest writes.
    ' TestHelper.ForceLocalBackend validates TbConfiguracionBackends is
    ' enabled, the path is local, and DAO can open it; it then caches the
    ' URL + password in m_BackendSandboxURL / m_BackendSandboxPassword and
    ' flips m_TestingMode = True. Failures raise Err 513.
    Call TestHelper.ForceLocalBackend(cfgError)
    If Len(cfgError) > 0 Then
        Call Sandbox_LogHarness("[Sandbox][EnsureSandboxReady] ForceLocalBackend blocked: " & cfgError)
        Err.Raise 513, "TestSandbox.EnsureSandboxReady", cfgError
    End If

    Call Sandbox_LogHarness("[Sandbox][EnsureSandboxReady] Sandbox routing asserted: " & m_BackendSandboxURL)

    ' getdb() now returns the asserted sandbox (m_TestingMode=True branch).
    Set EnsureSandboxReady = getdb()

    Exit Function

ErroresEnsure:
    Call Sandbox_LogHarness("[Sandbox][EnsureSandboxReady] ERROR " & Err.Number & " - " & Err.description)
    Err.Raise Err.Number, "TestSandbox.EnsureSandboxReady", Err.description
End Function

' ---
' ResetSandboxIfInvalid: Fuerza un rebuild del sandbox si es inválido.
' Útil para tests que requieren estado limpio o cuando se detecta corruption.
' El enfoque simplificado siempre hace build fresco (no hay validación compleja).
' ---
Public Sub ResetSandboxIfInvalid()
    Dim dbResult As DAO.Database
    
    On Error GoTo ErroresReset
    
    Call Sandbox_LogHarness("[Sandbox][ResetSandboxIfInvalid] Iniciando reset (build fresco)...")
    
    ' El enfoque simplificado no tiene validación compleja - siempre rebuild
    ' Esto es más simple y más robusto que el enfoque anterior
    Call Sandbox_LogHarness("[Sandbox][ResetSandboxIfInvalid] Forzando BuildSandbox...")
    Set dbResult = SandboxBuilder.BuildSandbox()
    Set m_SandboxDB = dbResult
    m_SandboxPath = SandboxBuilder.GetSandboxPath
    m_blnSandboxExists = True
    Set g_dbTestInstance = dbResult  ' Mantener g_dbTestInstance sincronizado
    
ErroresReset:
    Call Sandbox_LogHarness("[Sandbox][ResetSandboxIfInvalid] ERROR " & Err.Number & " - " & Err.description)
    Err.Raise Err.Number, "TestSandbox.ResetSandboxIfInvalid", Err.description
End Sub

' ==========================================================================
' PROVISIONING MÍNIMO — SOLO FILECOPY PARA INSPECCIÓN MANUAL
' ==========================================================================
' NO linked-table detection, NO sidecars, NO TransferDatabase, NO import.
' Solo copia el backend a la carpeta sandbox y retorna el path.
' ==========================================================================



' ---
' ProvisionSandbox_Minimo: Provisioning MÍNIMO para inspección manual.
'
'   POLÍTICA: Solo hace FileCopy del backend CONDOR a la carpeta sandbox local.
'   - NO detecta linked tables
'   - NO crea sidecars
'   - NO usa TransferDatabase ni import
'   - NO usa CondorError (solo Err.Raise nativo)
'   - Usa Debug.Print para logging (no depende de Canonical_Log_Trace)
'
'   RETORNA: String con el path completo del sandbox copy
'
'   USO DESDE Access VBA Immediate Window:
'     ? ProvisionSandbox_Minimo()
'     ' Debug.Print ProvisionSandbox_Minimo()
' ---
Public Function ProvisionSandbox_Minimo() As String
    Dim fso As New FileSystemObject
    Dim proyectoPath As String
    Dim sep As String
    Dim sandboxDir As String
    Dim sandboxPath As String
    Dim blnBackendExiste As Boolean
    
    On Error GoTo ErroresMinimo
    
    ' Logging de inicio
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo] ===== INICIO PROVISION MINIMO ====="
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo] Backend source: " & MINIMO_BACKEND_SOURCE
    
    ' --- Calcular rutas ---
    sep = "\"
    proyectoPath = CurrentProject.path
    If Right(proyectoPath, 1) <> sep Then proyectoPath = proyectoPath & sep
    
    sandboxDir = proyectoPath & MINIMO_SANDBOX_FOLDER
    sandboxPath = sandboxDir & sep & MINIMO_SANDBOX_FILENAME
    
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo] Sandbox dir: " & sandboxDir
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo] Sandbox path: " & sandboxPath
    
    ' --- Crear directorio de sandbox si no existe ---
    If Not fso.FolderExists(sandboxDir) Then
        fso.CreateFolder sandboxDir
        ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo] Directorio creado: " & sandboxDir
    End If
    
    ' --- Verificar que el backend existe ---
    blnBackendExiste = fso.FileExists(MINIMO_BACKEND_SOURCE)
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo] Backend existe: " & IIf(blnBackendExiste, "SI", "NO")
    
    If Not blnBackendExiste Then
        Err.Raise 20250, "TestSandbox.ProvisionSandbox_Minimo", _
            "Backend CONDOR no encontrado en: " & MINIMO_BACKEND_SOURCE & ". Verifique la conexión a la red."
    End If
    
    ' --- Eliminar sandbox anterior si existe ---
    If fso.FileExists(sandboxPath) Then
        fso.DeleteFile sandboxPath, True
        ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo] Sandbox previo eliminado."
    End If
    
    ' --- FILECOPY directo (sin linked tables, sin sidecars, sin import) ---
    fso.CopyFile MINIMO_BACKEND_SOURCE, sandboxPath, True
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo] FileCopy completado."
    
    ' --- Actualizar estado interno del módulo por compatibilidad ---
    m_SandboxPath = sandboxPath
    m_blnSandboxExists = True
    
    ' Logging de fin
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo] ===== PROVISION MINIMO FINALIZADO ====="
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo] Path del sandbox: " & sandboxPath
    
    ProvisionSandbox_Minimo = sandboxPath
    Exit Function
    
ErroresMinimo:
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo] ERROR " & err.Number & " - " & err.description
    Err.Raise Err.Number, "TestSandbox.ProvisionSandbox_Minimo", Err.description
End Function

' ---
' ProvisionSandbox_Minimo_Abrir: Provisioning mínimo Y retorna el DAO.Database abierto.
'   Igual que ProvisionSandbox_Minimo pero también abre el sandbox.
'   Útil para inspección programática.
'
'   RETORNA: DAO.Database del sandbox abierto (listo para consultar)
'   PARAMETRO dbOut: Database abierta por referencia (opcional)
' ---
Public Function ProvisionSandbox_Minimo_Abrir(Optional ByRef dbOut As DAO.Database = Nothing) As DAO.Database
    Dim dbResult As DAO.Database
    Dim sandboxPath As String
    Dim ws As DAO.Workspace
    
    On Error GoTo ErroresMinimoAbrir
    
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo_Abrir] ===== INICIO PROVISION MINIMO + ABRIR ====="
    
    ' Provisionar (copiar archivo)
    sandboxPath = ProvisionSandbox_Minimo()
    
    ' Asegurar workspace aislado
    Set ws = getWorkspace()
    
    ' Abrir el sandbox con password
    Set dbResult = ws.OpenDatabase(sandboxPath, False, False, ";pwd=" & GetPasswordDB())
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo_Abrir] Sandbox abierto como DAO.Database."
    
    ' Devolver por parámetro ByRef si se solicitó
    Set dbOut = dbResult
    
    ' Actualizar estado interno
    Set m_SandboxDB = dbResult
    Set m_SandboxWS = ws
    
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo_Abrir] ===== PROVISION + ABRIERTO FINALIZADO ====="
    
    Set ProvisionSandbox_Minimo_Abrir = dbResult
    Exit Function
    
ErroresMinimoAbrir:
    ' Debug.Print "[TestSandbox][ProvisionSandbox_Minimo_Abrir] ERROR " & err.Number & " - " & err.description
    Err.Raise Err.Number, "TestSandbox.ProvisionSandbox_Minimo_Abrir", Err.description
End Function


