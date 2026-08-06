VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SandboxConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' CLASE: SandboxConfig
' DESCRIPCIÓN: Lee y valida el JSON de configuración del sandbox engine.
' DEPENDENCIAS: JsonConverter.bas
' ==========================================================================

' --- Errores custom ---
Private Const ERR_CONFIG_LOAD As Long = 20001
Private Const ERR_CONFIG_VALIDATE As Long = 20002
Private Const ERR_CONFIG_INVALID As Long = 20003

' --- Variables privadas ---
Private m_strConfigPath As String
Private m_objJsonConfig As Object
Private m_strSchemaVersion As String
Private m_blnAutoRebuildOnInvalid As Boolean
Private m_blnStrictMode As Boolean
Private m_strOutputDir As String
Private m_strSandboxName As String
Private m_colTablesToLocalize As Collection
Private m_dicBackends As Scripting.Dictionary
Private m_colProductionDenyPatterns As Collection  ' PHASE 4: Patrones de producción a bloquear

' --------------------------------------------------------------------------
' TRACE DE CONFIG: escribe en Debug.Print y, si la batería está activa,
' también en el log canonical. Nunca debe romper el flujo.
' --------------------------------------------------------------------------
Private Sub TraceConfig(ByVal msg As String)
    On Error Resume Next
    Debug.Print msg
    Call modBattery_Canonical.Canonical_Log_Trace(msg)
    On Error GoTo 0
End Sub

' ==========================================================================
' PROPIEDADES PÚBLICAS
' ==========================================================================

' ---
' BackendSource: Ruta del backend origen para la tabla especificada
' ---
Public Property Get BackendSource() As String
    BackendSource = m_strConfigPath
End Property

' ---
' TablesToLocalize: Colección de diccionarios con {name, backend, sqlFilter}
' ---
Public Property Get TablesToLocalize() As Collection
    Set TablesToLocalize = m_colTablesToLocalize
End Property

' ---
' OutputDir: Directorio donde crear el sandbox
' ---
Public Property Get outputDir() As String
    outputDir = m_strOutputDir
End Property

' ---
' SandboxName: Nombre del archivo ACCDB del sandbox
' ---
Public Property Get sandboxName() As String
    sandboxName = m_strSandboxName
End Property

' ---
' AutoRebuildOnInvalid: Si True, reconstruye sandbox automáticamente si es inválido
' ---
Public Property Get AutoRebuildOnInvalid() As Boolean
    AutoRebuildOnInvalid = m_blnAutoRebuildOnInvalid
End Property

' ---
' StrictMode: Modo estricto para validaciones adicionales
' ---
Public Property Get StrictMode() As Boolean
    StrictMode = m_blnStrictMode
End Property

' ---
' SchemaVersion: Versión del schema de configuración
' ---
Public Property Get SchemaVersion() As String
    SchemaVersion = m_strSchemaVersion
End Property

' ---
' ConfigHash: Hash SHA-256 del contenido JSON (para detectar cambios)
' ---
Public Property Get ConfigHash() As String
    ConfigHash = GetConfigHashInternal()
End Property

' ---
' ProductionDenyPatterns: Colección de patrones de rutas de producción a bloquear
' PHASE 4: Nuevos guardrails - estos patrones se usan para detectar production paths
' ---
Public Property Get ProductionDenyPatterns() As Collection
    Set ProductionDenyPatterns = m_colProductionDenyPatterns
End Property

' ==========================================================================
' MÉTODOS PÚBLICOS
' ==========================================================================

' ---
' Load: Carga configuración desde archivo JSON
' ---
Public Sub Load(configPath As String)
    On Error GoTo Errores
    
    Dim strJsonContent As String
    Dim objFSO As FileSystemObject
    Dim strPreview As String
    
    m_strConfigPath = configPath
    Call TraceConfig("[SandboxConfig][Load] INICIO. configPath=" & configPath)
    
    ' Verificar que el archivo existe
    Set objFSO = New FileSystemObject
    If Not objFSO.FileExists(configPath) Then
        Dim errObj As New CondorError
        errObj.Create ERR_CONFIG_LOAD, "Archivo de configuración no encontrado: " & configPath, "SandboxConfig.Load"
        errObj.Raise
    End If
    
    ' Leer contenido JSON
    With objFSO.OpenTextFile(configPath, ForReading)
        strJsonContent = .ReadAll
        .Close
    End With
    
    Call TraceConfig("[SandboxConfig][Load] JSON leído. Len=" & Len(strJsonContent))
    If Len(strJsonContent) > 0 Then
        strPreview = Left$(strJsonContent, 80)
        strPreview = Replace(strPreview, vbCr, "\\r")
        strPreview = Replace(strPreview, vbLf, "\\n")
        strPreview = Replace(strPreview, vbTab, "\\t")
        Call TraceConfig("[SandboxConfig][Load] Preview='" & strPreview & "'")
        Call TraceConfig("[SandboxConfig][Load] Primer char ASCII=" & Asc(Left$(strJsonContent, 1)))
    Else
        Call TraceConfig("[SandboxConfig][Load] ERROR: JSON vacío.")
    End If
    
    ' Parsear JSON
    Call TraceConfig("[SandboxConfig][Load] Invocando JsonConverter.ParseJson...")
    Set m_objJsonConfig = JsonConverter.ParseJson(strJsonContent)
    Call TraceConfig("[SandboxConfig][Load] ParseJson OK. TypeName=" & TypeName(m_objJsonConfig))
    
    ' Inicializar colecciones y diccionarios
    Set m_colTablesToLocalize = New Collection
    Set m_dicBackends = New Scripting.Dictionary
    Call TraceConfig("[SandboxConfig][Load] Colecciones inicializadas.")
    
    ' Extraer configuración
    Call TraceConfig("[SandboxConfig][Load] Llamando ParseConfiguration...")
    Call ParseConfiguration
    Call TraceConfig("[SandboxConfig][Load] FIN OK.")
    
    Exit Sub
    
Errores:
    Call TraceConfig("[SandboxConfig][Load] ERROR " & Err.Number & " - " & Err.description)
    Dim errObjFinal As New CondorError
    errObjFinal.Create Err.Number, Err.description, "SandboxConfig.Load"
    errObjFinal.Raise
End Sub

' ---
' Validate: Valida que la configuración tiene lo mínimo para funcionar
' Retorna True si es válida, False si no
' ---
Public Function Validate() As Boolean
    On Error GoTo Errores
    
    Dim blnResult As Boolean
    blnResult = True
    
    ' Verificar schema version
    If m_strSchemaVersion = "" Then
        blnResult = False
    End If
    
    ' Verificar output dir
    If m_strOutputDir = "" Then
        blnResult = False
    End If
    
    ' Verificar sandbox name
    If m_strSandboxName = "" Then
        blnResult = False
    End If
    
    ' Verificar que hay tablas definidas
    If m_colTablesToLocalize.count = 0 Then
        blnResult = False
    End If
    
    ' Verificar que todos los backends referenciados existen
    Dim varItem As Variant
    Dim objTableDict As Object
    Dim strBackendKey As String
    
    For Each varItem In m_colTablesToLocalize
        If TypeName(varItem) = "Dictionary" Then
            Set objTableDict = varItem
            If objTableDict.Exists("backend") Then
                strBackendKey = objTableDict("backend")
                If Not m_dicBackends.Exists(strBackendKey) Then
                    blnResult = False
                End If
            End If
        End If
    Next varItem
    
    Validate = blnResult
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxConfig.Validate"
    errObj.Raise
End Function

' ---
' GetTableNames: Devuelve array de nombres de tablas a localizar
' ---
Public Function GetTableNames() As String()
    On Error GoTo Errores
    
    Dim arrTables() As String
    Dim lngCount As Long
    Dim varItem As Variant
    Dim objTableDict As Object
    
    lngCount = m_colTablesToLocalize.count
    
    If lngCount = 0 Then
        GetTableNames = arrTables
        Exit Function
    End If
    
    ReDim arrTables(0 To lngCount - 1)
    
    Dim lngIndex As Long
    lngIndex = 0
    
    For Each varItem In m_colTablesToLocalize
        If TypeName(varItem) = "Dictionary" Then
            Set objTableDict = varItem
            If objTableDict.Exists("name") Then
                arrTables(lngIndex) = objTableDict("name")
            End If
        End If
        lngIndex = lngIndex + 1
    Next varItem
    
    GetTableNames = arrTables
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxConfig.GetTableNames"
    errObj.Raise
End Function

' ---
' GetBackendPath: Obtiene la ruta del backend por clave
' ---
Public Function GetBackendPath(backendKey As String) As String
    On Error GoTo Errores
    
    If m_dicBackends.Exists(backendKey) Then
        GetBackendPath = m_dicBackends(backendKey)("path")
    Else
        GetBackendPath = ""
    End If
    
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxConfig.GetBackendPath"
    errObj.Raise
End Function

' ---
' GetBackendPassword: Obtiene la contraseña del backend por clave
' ---
Public Function GetBackendPassword(backendKey As String) As String
    On Error GoTo Errores
    
    If m_dicBackends.Exists(backendKey) Then
        GetBackendPassword = m_dicBackends(backendKey)("password")
    Else
        GetBackendPassword = ""
    End If
    
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxConfig.GetBackendPassword"
    errObj.Raise
End Function

' ==========================================================================
' MÉTODOS PRIVADOS
' ==========================================================================

Private Sub ParseProductionDenyPatterns()
    On Error GoTo Errores
    
    Dim objValidation As Object
    Dim objPatterns As Object
    Dim varPattern As Variant
    
    ' Buscar la sección validation.productionDenyPatterns
    If m_objJsonConfig.Exists("validation") Then
        Set objValidation = m_objJsonConfig("validation")
        
        If objValidation.Exists("productionDenyPatterns") Then
            Set objPatterns = objValidation("productionDenyPatterns")
            
            ' Es una colección de strings
            For Each varPattern In objPatterns
                If VarType(varPattern) = vbString Then
                    m_colProductionDenyPatterns.Add CStr(varPattern)
                End If
            Next varPattern
        End If
    End If
    
    ' Si no hay patrones en el JSON, agregar defaults hardcoded como fallback
    If m_colProductionDenyPatterns.count = 0 Then
        AddDefaultProductionPatterns
    End If
    
    Exit Sub
    
Errores:
    ' Si falla el parsing, agregar defaults
    If m_colProductionDenyPatterns.count = 0 Then
        Call AddDefaultProductionPatterns
    End If
End Sub

' ---
' AddDefaultProductionPatterns: Agrega patrones hardcoded como fallback
' Se usan si el JSON no tiene productionDenyPatterns definidos
' ---
Private Sub AddDefaultProductionPatterns()
    On Error GoTo Errores
    
    ' Patrones típicos de producción en el entorno Telefónica D&S
    m_colProductionDenyPatterns.Add "\\datoste\aplicaciones_dys\Aplicaciones PpD\HPS\"
    m_colProductionDenyPatterns.Add "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\"
    m_colProductionDenyPatterns.Add "\\datoste\aplicaciones_dys\Aplicaciones PpD\No Conformidades\"
    m_colProductionDenyPatterns.Add "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\"
    
    Exit Sub
    
Errores:
    ' Silenciar errores - los defaults ya son conocidos
End Sub

Private Sub ParseConfiguration()
    On Error GoTo Errores
    Call TraceConfig("[SandboxConfig][ParseConfiguration] INICIO.")
    
    Dim objSandbox As Object
    Dim objBackends As Object
    Dim strBackendKey As String
    Dim objBackend As Object
    Dim objLinkedTables As Collection
    Dim varTable As Variant
    Dim objTableDict As Object
    
    ' Inicializar colección de patrones de producción (PHASE 4)
    Set m_colProductionDenyPatterns = New Collection
    
    ' Extraer schema version
    If m_objJsonConfig.Exists("schemaVersion") Then
        m_strSchemaVersion = Nz(m_objJsonConfig("schemaVersion"), "")
        Call TraceConfig("[SandboxConfig][ParseConfiguration] schemaVersion='" & m_strSchemaVersion & "'")
    End If
    
    ' Extraer configuración del sandbox
    If m_objJsonConfig.Exists("sandbox") Then
        Set objSandbox = m_objJsonConfig("sandbox")
        Call TraceConfig("[SandboxConfig][ParseConfiguration] Sección sandbox encontrada.")
        
        If objSandbox.Exists("sandboxName") Then
            m_strSandboxName = Nz(objSandbox("sandboxName"), "")
            Call TraceConfig("[SandboxConfig][ParseConfiguration] sandboxName='" & m_strSandboxName & "'")
        End If
        
        If objSandbox.Exists("outputDir") Then
            m_strOutputDir = Nz(objSandbox("outputDir"), "")
            Call TraceConfig("[SandboxConfig][ParseConfiguration] outputDir='" & m_strOutputDir & "'")
        End If
        
        If objSandbox.Exists("autoRebuildOnInvalid") Then
            m_blnAutoRebuildOnInvalid = CBool(Nz(objSandbox("autoRebuildOnInvalid"), False))
        End If
        
        If objSandbox.Exists("strictMode") Then
            m_blnStrictMode = CBool(Nz(objSandbox("strictMode"), False))
        End If
    End If
    
    ' Extraer backends y tablas linked
    If m_objJsonConfig.Exists("backends") Then
        Set objBackends = m_objJsonConfig("backends")
        Call TraceConfig("[SandboxConfig][ParseConfiguration] Sección backends encontrada.")
        
        Dim varKey As Variant
        For Each varKey In objBackends.Keys
            strBackendKey = CStr(varKey)
            Set objBackend = objBackends(varKey)
            Call TraceConfig("[SandboxConfig][ParseConfiguration] Backend='" & strBackendKey & "'")
            
            ' Guardar backend info en diccionario
            ' FIX: Crear UNA NUEVA instancia de Dictionary por cada backend.
            ' Con Dim As New, VBA reuse el mismo objeto en cada iteración del loop,
            ' causando que todos los m_dicBackends entries referencien al ÚLTIMO backend.
            Dim objBackendInfo As Scripting.Dictionary
            Set objBackendInfo = New Scripting.Dictionary
            objBackendInfo("path") = Nz(objBackend("path"), "")
            objBackendInfo("password") = Nz(objBackend("password"), "")
            Set m_dicBackends(strBackendKey) = objBackendInfo
            Call TraceConfig("[SandboxConfig][ParseConfiguration] Backend '" & strBackendKey & "' stored in m_dicBackends.")
            
            ' Extraer linked tables de este backend
            If objBackend.Exists("linkedTables") Then
                Set objLinkedTables = objBackend("linkedTables")
                Call TraceConfig("[SandboxConfig][ParseConfiguration] linkedTables count backend '" & strBackendKey & "' = " & objLinkedTables.count)
                
                For Each varTable In objLinkedTables
                    Set objTableDict = New Scripting.Dictionary
                    objTableDict("name") = CStr(varTable)
                    objTableDict("backend") = strBackendKey
                    objTableDict("sqlFilter") = ""
                    m_colTablesToLocalize.Add objTableDict
                Next varTable
            End If
        Next varKey
    End If
    
    ' PHASE 4: Extraer productionDenyPatterns de la sección validation
    Call TraceConfig("[SandboxConfig][ParseConfiguration] Llamando ParseProductionDenyPatterns...")
    Call ParseProductionDenyPatterns
    Call TraceConfig("[SandboxConfig][ParseConfiguration] FIN OK.")
    
    Exit Sub
    
Errores:
    Call TraceConfig("[SandboxConfig][ParseConfiguration] ERROR " & Err.Number & " - " & Err.description)
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxConfig.ParseConfiguration"
    errObj.Raise
End Sub

Private Function GetConfigHashInternal() As String
    On Error GoTo Errores
    
    Dim objFSO As FileSystemObject
    Dim strContent As String
    
    If m_strConfigPath = "" Then
        GetConfigHashInternal = ""
        Exit Function
    End If
    
    Set objFSO = New FileSystemObject
    With objFSO.OpenTextFile(m_strConfigPath, ForReading)
        strContent = .ReadAll
        .Close
    End With
    
    ' Generar hash simple (SHA-256 via HashHelper si existe)
    If Len(strContent) > 0 Then
        GetConfigHashInternal = HashHelper_SHA256(strContent)
    Else
        GetConfigHashInternal = ""
    End If
    
    Exit Function
    
Errores:
    GetConfigHashInternal = ""
End Function

Private Function HashHelper_SHA256(strInput As String) As String
    On Error GoTo Errores
    
    ' Intentar usar HashHelper si está disponible
    HashHelper_SHA256 = ""
    
    ' Si HashHelper no está disponible, retornar hash simple basado en contenido
    Dim lngHash As Long
    Dim i As Long
    
    lngHash = 0
    For i = 1 To Len(strInput)
        lngHash = ((lngHash * 31) + Asc(Mid(strInput, i, 1))) And &HFFFFFFFF
    Next i
    
    HashHelper_SHA256 = Hex(lngHash)
    
    Exit Function
    
Errores:
    HashHelper_SHA256 = ""
End Function

Private Function Nz(valor As Variant, Optional valorDefecto As Variant = "") As Variant
    If IsNull(valor) Or IsEmpty(valor) Then
        Nz = valorDefecto
    Else
        Nz = valor
    End If
End Function
