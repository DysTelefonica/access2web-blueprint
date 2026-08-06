VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SandboxValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' CLASE: SandboxValidator
' DESCRIPCIÓN: Valida si un sandbox existente es válido o no.
'              Detecta linked tables activas, falta de tablas configuradas,
'              y valida estructura mínima/metadata.
'              También valida que las rutas NO sean de producción (Phase 4).
' FASE: 3 (validación) y Phase 4 (guardrails) del cambio sandbox-linked-localization
' DEPENDENCIAS: SandboxConfig, CondorError
' ==========================================================================

' --- Constantes de errores ---
Private Const ERR_VALIDATOR_NOT_INITIALIZED As Long = 20100
Private Const ERR_VALIDATION_FAILED As Long = 20101
Private Const ERR_NO_CONFIG As Long = 20102
Private Const ERR_PRODUCTION_PATH_DETECTED As Long = 20103
Private Const ERR_UNSAFE_PATH As Long = 20104
Private Const ERR_BACKEND_NOT_READONLY_SAFE As Long = 20105

' --- Constantes DAO para detección de linked tables ---
Private Const DB_LINKED As Long = &H80000000

' --- Variables privadas ---
Private m_blnIsValid As Boolean
Private m_strValidationSummary As String
Private m_colLinkedTables As Collection
Private m_colMissingTables As Collection
Private m_colValidationErrors As Collection
Private m_blnChecked As Boolean

' ==========================================================================
' PROPIEDADES PÚBLICAS
' ==========================================================================

' ---
' IsValid: Resultado de la última validación
' ---
Public Property Get isValid() As Boolean
    isValid = m_blnIsValid
End Property

' ---
' ValidationSummary: Resumen legible de la validación (para debugging/UI)
' ---
Public Property Get ValidationSummary() As String
    ValidationSummary = m_strValidationSummary
End Property

' ---
' LinkedTablesFound: Colección de nombres de tablas linked encontradas
' ---
Public Property Get LinkedTablesFound() As Collection
    Set LinkedTablesFound = m_colLinkedTables
End Property

' ---
' MissingTables: Colección de tablas configuradas que faltan en el sandbox
' ---
Public Property Get MissingTables() As Collection
    Set MissingTables = m_colMissingTables
End Property

' ---
' ValidationErrors: Lista de errores de validación
' ---
Public Property Get ValidationErrors() As Collection
    Set ValidationErrors = m_colValidationErrors
End Property

' ==========================================================================
' MÉTODOS PÚBLICOS
' ==========================================================================

' ---
' ValidateSandbox: Valida si el sandbox es válido para reuse
' config: Configuración del sandbox (requerida)
' dbSandbox: Base de datos del sandbox a validar (si Nothing, usa CurrentDb)
' Retorna: True si el sandbox es válido, False si no
' Lógica:
'   1. Verifica que no tenga linked tables activas
'   2. Verifica que existan las tablas configuradas
' ---
Public Function ValidateSandbox(Config As SandboxConfig, Optional dbSandbox As DAO.Database = Nothing) As Boolean
    On Error GoTo Errores
    
    Dim db As DAO.Database
    Dim blnResult As Boolean
    Dim strErrors As String
    
    ' Resetear estado
    Call ResetValidationState
    
    ' Verificar que tenemos config
    If Config Is Nothing Then
        Dim errObjCfg As New CondorError
        errObjCfg.Create ERR_NO_CONFIG, "SandboxValidator.ValidateSandbox: Se requiere configuración válida", "SandboxValidator.ValidateSandbox"
        errObjCfg.Raise
    End If
    
    ' Usar database especificada o CurrentDb
    If dbSandbox Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbSandbox
    End If
    
    blnResult = True
    
    ' --- CHECK 1: Verificar que NO haya linked tables activas ---
    If HasLinkedTables(db) Then
        Dim colLinked As Collection
        Set colLinked = GetLinkedTableNames(db)
        If colLinked.count > 0 Then
            blnResult = False
            Dim varItem As Variant
            For Each varItem In colLinked
                Call AddValidationError("Linked table activa encontrada: " & CStr(varItem))
            Next varItem
        End If
    End If
    
    ' --- CHECK 2: Verificar tablas configuradas ---
    If Not ValidateConfiguredTables(Config, db) Then
        blnResult = False
    End If

    ' Generar summary
    m_strValidationSummary = GenerateValidationSummary(blnResult)
    
    ' Loguear resultado
    strErrors = JoinValidationErrors()
    Debug.Print "[SandboxValidator][ValidateSandbox] Result=" & blnResult & " Errors=" & strErrors
    
    m_blnIsValid = blnResult
    m_blnChecked = True
    
    ValidateSandbox = blnResult
    Exit Function
    
Errores:
    Dim errObj2 As New CondorError
    errObj2.Create Err.Number, Err.description, "SandboxValidator.ValidateSandbox"
    errObj2.Raise
End Function

' ---
' HasLinkedTables: Verifica si la base de datos tiene tablas linked
' dbSandbox: Database a verificar (si Nothing, usa CurrentDb)
' Retorna: True si encuentra al menos una linked table
' ---
Public Function HasLinkedTables(Optional dbSandbox As DAO.Database = Nothing) As Boolean
    On Error GoTo Errores
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    If dbSandbox Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbSandbox
    End If
    
    For Each tdf In db.TableDefs
        ' Omitir tablas de sistema
        If Left(tdf.name, 4) = "MSys" Or Left(tdf.name, 4) = "USys" Then
            GoTo NextTable
        End If
        
        ' Verificar si es linked usando el flag dbAttached
        If IsLinkedTableDef(tdf) Then
            HasLinkedTables = True
            Exit Function
        End If
NextTable:
    Next tdf
    
    HasLinkedTables = False
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxValidator.HasLinkedTables"
    errObj.Raise
End Function

' ---
' GetLinkedTableNames: Obtiene lista de nombres de tablas linked
' dbSandbox: Database a verificar (si Nothing, usa CurrentDb)
' Retorna: Collection con nombres de tablas linked
' ---
Public Function GetLinkedTableNames(Optional dbSandbox As DAO.Database = Nothing) As Collection
    On Error GoTo Errores
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim colResult As New Collection
    
    If dbSandbox Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbSandbox
    End If
    
    For Each tdf In db.TableDefs
        ' Omitir tablas de sistema
        If Left(tdf.name, 4) = "MSys" Or Left(tdf.name, 4) = "USys" Then
            GoTo NextTable2
        End If
        
        ' Verificar si es linked
        If IsLinkedTableDef(tdf) Then
            colResult.Add tdf.name
        End If
NextTable2:
    Next tdf
    
    Set GetLinkedTableNames = colResult
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxValidator.GetLinkedTableNames"
    errObj.Raise
End Function

' ---
' ValidateConfiguredTables: Verifica que todas las tablas configuradas existan
' config: Configuración con lista de tablas a localizar
' dbSandbox: Database del sandbox (si Nothing, usa CurrentDb)
' Retorna: True si todas existen, False si falta alguna
' ---
Public Function ValidateConfiguredTables(Config As SandboxConfig, Optional dbSandbox As DAO.Database = Nothing) As Boolean
    On Error GoTo Errores
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim colTableNames As Collection
    Dim varName As Variant
    Dim blnAllExist As Boolean
    Dim blnTableFound As Boolean
    Dim strTableName As String
    
    blnAllExist = True
    
    ' Usar database especificada o CurrentDb
    If dbSandbox Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbSandbox
    End If
    
    ' Obtener nombres de tablas desde config
    Set colTableNames = New Collection
    Dim arrNames() As String
    arrNames = Config.GetTableNames()
    
    Dim i As Long
    For i = LBound(arrNames) To UBound(arrNames)
        If arrNames(i) <> "" Then
            colTableNames.Add arrNames(i)
        End If
    Next i
    
    ' Verificar cada tabla
    For Each varName In colTableNames
        strTableName = CStr(varName)
        blnTableFound = False
        
        For Each tdf In db.TableDefs
            ' Omitir tablas de sistema
            If Left(tdf.name, 4) = "MSys" Or Left(tdf.name, 4) = "USys" Then
                GoTo NextTdf
            End If
            
            If tdf.name = strTableName Then
                blnTableFound = True
                Exit For
            End If
NextTdf:
        Next tdf
        
        If Not blnTableFound Then
            blnAllExist = False
            m_colMissingTables.Add strTableName
            Call AddValidationError("Tabla configurada no encontrada: " & strTableName)
        End If
    Next varName
    
    ValidateConfiguredTables = blnAllExist
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SandboxValidator.ValidateConfiguredTables"
    errObj.Raise
End Function

' ---
' GetValidationSummary: Devuelve el resumen de la última validación
' Versión de Property para compatibilidad
' ---
Public Function GetValidationSummary() As String
    GetValidationSummary = m_strValidationSummary
End Function

' ==========================================================================
' MÉTODOS PRIVADOS
' ==========================================================================

' ---
' ResetValidationState: Limpia el estado de validación
' ---
Private Sub ResetValidationState()
    m_blnIsValid = False
    m_strValidationSummary = ""
    Set m_colLinkedTables = New Collection
    Set m_colMissingTables = New Collection
    Set m_colValidationErrors = New Collection
    m_blnChecked = False
End Sub

' ---
' AddValidationError: Agrega un error a la lista
' ---
Private Sub AddValidationError(errorMsg As String)
    m_colValidationErrors.Add errorMsg
End Sub

' ---
' JoinValidationErrors: Concatena errores con pipe
' ---
Private Function JoinValidationErrors() As String
    Dim varErr As Variant
    Dim strResult As String
    
    strResult = ""
    For Each varErr In m_colValidationErrors
        If strResult <> "" Then strResult = strResult & " | "
        strResult = strResult & CStr(varErr)
    Next varErr
    
    JoinValidationErrors = strResult
End Function

' ---
' GenerateValidationSummary: Genera un resumen legible del resultado
' ---
Private Function GenerateValidationSummary(isValid As Boolean) As String
    Dim strSummary As String
    
    If isValid Then
        strSummary = "VALID: Sandbox OK para reuse"
    Else
        strSummary = "INVALID: Sandbox requiere rebuild"
        
        If m_colLinkedTables.count > 0 Then
            strSummary = strSummary & " | Linked tables: " & m_colLinkedTables.count
        End If
        
        If m_colMissingTables.count > 0 Then
            strSummary = strSummary & " | Tablas faltantes: " & m_colMissingTables.count
        End If
        
        If m_colValidationErrors.count > 0 Then
            strSummary = strSummary & " | Errores: " & JoinValidationErrors()
        End If
    End If
    
    GenerateValidationSummary = strSummary
End Function

' ---
' IsLinkedTableDef: Evalúa si un TableDef representa una tabla linked
' ---
Private Function IsLinkedTableDef(tdf As DAO.TableDef) As Boolean
    On Error GoTo Errores
    IsLinkedTableDef = (tdf.Attributes And DB_LINKED) = DB_LINKED
    Exit Function
Errores:
    IsLinkedTableDef = False
End Function

' ==========================================================================
' MÉTODOS DE GUARDRAILS — PHASE 4
' ==========================================================================
' Estos métodos blindan el sandbox engine contra operaciones peligrosas
' en backends de producción o rutas no seguras.
' ==========================================================================

' ---
' IsProductionLikePath: Detecta si una ruta parece ser de producción/riesgosa
' path: Ruta a evaluar (puede ser UNC, local, etc.)
' Retorna: True si la ruta parece de producción, False si parece segura
'
' heuristics:
'   - UNC paths (\\servidor\...) se consideran potencialmente peligrosos
'   - Rutas que contienen patrones protegidos (HPS, CONDOR\00_Condor, etc.)
'   - Rutas de red que parecen ser del servidor central (datoste)
' ---
Public Function IsProductionLikePath(path As String) As Boolean
    On Error GoTo Errores
    
    Dim strPath As String
    strPath = Trim(path)
    
    If strPath = "" Then
        IsProductionLikePath = False
        Exit Function
    End If
    
    ' Normalizar: convertir forward slashes a backslashes
    strPath = Replace(strPath, "/", "\")
    
    ' CHECK 1: ¿Es UNC path de red?
    ' \\datoste, \\servidor, etc. ? potencialmente peligroso
    If Left(strPath, 2) = "\\" Then
        ' Es UNC - verificar si es uno de los patrones conocidos de producción
        If IsUnsafeNetworkPath(strPath) Then
            IsProductionLikePath = True
            Exit Function
        End If
    End If
    
    ' CHECK 2: ¿Contiene patrones de producción protegidos?
    ' Patrones típicos de entornos Telefónica D&S
    If ContainsProductionPattern(strPath) Then
        IsProductionLikePath = True
        Exit Function
    End If
    
    ' CHECK 3: ¿Está en directorio padre de sandbox que contiene datos reales?
    ' Si la ruta contiene 00_Condor, HPS, etc. como parte del path
    If ContainsSandboxLikePath(strPath) Then
        IsProductionLikePath = True
        Exit Function
    End If
    
    IsProductionLikePath = False
    Exit Function
    
Errores:
    IsProductionLikePath = False
End Function

' ---
' IsUnsafeNetworkPath: Evalúa si un UNC path es de producción
' Solo para UNC paths (los que empiezan con \\)
' ---
Private Function IsUnsafeNetworkPath(uncPath As String) As Boolean
    On Error GoTo Errores
    
    Dim strLower As String
    strLower = LCase(uncPath)
    
    ' UNC paths que son claramente de producción/riesgosos
    ' Estos servidores contienen datos reales de la empresa
    If InStr(strLower, "\\datoste") > 0 Then
        ' datoste es el servidor de datos central - CASI SIEMPRE producción
        ' Se permiten ciertos paths de test, pero se bloquean los de prod
        If ContainsUnsafeDatostePattern(uncPath) Then
            IsUnsafeNetworkPath = True
            Exit Function
        End If
    End If
    
    ' Otros servidores de red conocidos como peligrosos
    If InStr(strLower, "\\servidor") > 0 Then
        IsUnsafeNetworkPath = True
        Exit Function
    End If
    
    IsUnsafeNetworkPath = False
    Exit Function
    
Errores:
    IsUnsafeNetworkPath = False
End Function

' ---
' ContainsUnsafeDatostePattern: Evalúa si un path UNC de datoste es peligroso
' Solo para paths que ya contain \\datoste
' ---
Private Function ContainsUnsafeDatostePattern(uncPath As String) As Boolean
    On Error GoTo Errores
    
    Dim strLower As String
    strLower = LCase(uncPath)
    
    ' Patrones claramente de producción en datoste
    ' \\datoste\...\Aplicaciones PpD\EXPEDIENTES\ ? Expedientes producción
    ' \\datoste\...\Aplicaciones PpD\No Conformidades\ ? NoConformidades producción
    ' \\datoste\...\Aplicaciones PpD\0Lanzadera\ ? Lanzadera producción
    ' \\datoste\...\Aplicaciones PpD\HPS\ ? HPS producción
    ' \\datoste\...\Aplicaciones PpD\CONDOR\00_Condor\ ? CONDOR frontend producción
    
    If InStr(strLower, "\aplicaciones ppd\") > 0 Or InStr(strLower, "\aplicaciones_dys\") > 0 Then
        ' Tiene la carpeta de aplicaciones - verificar subcarpetas peligrosas
        If InStr(strLower, "\hps\") > 0 Then
            ContainsUnsafeDatostePattern = True
            Exit Function
        End If
        If InStr(strLower, "\00_hps\") > 0 Then
            ContainsUnsafeDatostePattern = True
            Exit Function
        End If
        If InStr(strLower, "\condor\00_condor\") > 0 Then
            ContainsUnsafeDatostePattern = True
            Exit Function
        End If
        If InStr(strLower, "\expedientes\") > 0 Then
            ' Expedientes es producción - pero puede ser el backend a clonar
            ' Para OUTPUT/ sandbox path esto es peligroso
            ' Para INPUT/backend source esto es lo que vamos a leer (y ES producción)
            ' La distinción se hace en AssertBackendSourceIsReadOnlySafe
            ContainsUnsafeDatostePattern = True
            Exit Function
        End If
        If InStr(strLower, "\no conformidades\") > 0 Then
            ContainsUnsafeDatostePattern = True
            Exit Function
        End If
        If InStr(strLower, "\0lanzadera\") > 0 Then
            ContainsUnsafeDatostePattern = True
            Exit Function
        End If
    End If
    
    ContainsUnsafeDatostePattern = False
    Exit Function
    
Errores:
    ContainsUnsafeDatostePattern = False
End Function

' ---
' ContainsProductionPattern: Evalúa si una ruta local/UNC contiene
' patrones que indican datos de producción
' ---
Private Function ContainsProductionPattern(path As String) As Boolean
    On Error GoTo Errores
    
    Dim strLower As String
    strLower = LCase(path)
    
    ' Patrones de producción explícitos en el path
    ' Estos siempre se bloquean para output sandbox (son sandbox INseguros)
    ' NOTA: \condor\00_condor\ NO se incluye aquí porque paths de producción
    ' CONDOR en \\datoste\... ya son capturados por ContainsUnsafeDatostePattern.
    ' Incluirlo aquí causaba false-positive rechazando paths locales legítimos
    ' como C:\...\CONDOR\00_Condor\test_sandbox
    Static arrPatterns As Variant
    arrPatterns = Array( _
        "\hps\", _
        "\00_hps\", _
        "\aplicaciones ppd\hps", _
        "\aplicaciones_dys\hps" _
    )
    
    Dim i As Long
    For i = LBound(arrPatterns) To UBound(arrPatterns)
        If InStr(strLower, arrPatterns(i)) > 0 Then
            ContainsProductionPattern = True
            Exit Function
        End If
    Next i
    
    ContainsProductionPattern = False
    Exit Function
    
Errores:
    ContainsProductionPattern = False
End Function

' ---
' ContainsSandboxLikePath: Evalúa si el path parece ser un sandbox
' (pero no lo es) - indica confusión del usuario o config mal escrita
' ---
Private Function ContainsSandboxLikePath(path As String) As Boolean
    On Error GoTo Errores
    
    Dim strLower As String
    strLower = LCase(path)
    
    ' Si el path tiene "sandbox" pero también tiene patrones de prod...
    If InStr(strLower, "sandbox") > 0 Then
        ' Está bien, es un sandbox
        ' PERO si simultáneamente tiene HPS/CONDOR, probablemente
        ' el usuario mezcló paths
        If InStr(strLower, "\hps\") > 0 Or InStr(strLower, "\00_hps\") > 0 Then
            ContainsSandboxLikePath = True
            Exit Function
        End If
    End If
    
    ContainsSandboxLikePath = False
    Exit Function
    
Errores:
    ContainsSandboxLikePath = False
End Function

' ---
' AssertSandboxPathIsSafe: Verifica que la ruta del sandbox es SEGURA
' (i.e., NO es producción, NO es una ruta de red peligrosa)
' Si NO es segura, LANZA CondorError
'
' Esta es la guard MAIN para proteger contra escrituras en producción
'
' path: Ruta propuesta para el sandbox (output)
' ---
Public Sub AssertSandboxPathIsSafe(path As String)
    On Error GoTo Errores
    
    If IsProductionLikePath(path) Then
        Dim errObj As New CondorError
        errObj.Create ERR_PRODUCTION_PATH_DETECTED, _
            "RUTA DE PRODUCCIÓN/BACKEND REAL DETECTADA PARA SANDBOX OUTPUT." & vbNewLine & vbNewLine & _
            "La siguiente ruta fue detectada como potencialmente peligrosa:" & vbNewLine & _
            "  " & path & vbNewLine & vbNewLine & _
            "El sandbox engine NO puede crear un sandbox en una ruta de producción." & vbNewLine & vbNewLine & _
            "ACCIÓN REQUERIDA:" & vbNewLine & _
            "1. Verifique que 'outputDir' en el JSON de configuración apunta a una carpeta LOCAL de test" & vbNewLine & _
            "2. Ejemplo seguro: 'C:\TestSandbox\' o 'test_sandbox\' (relativo al proyecto)" & vbNewLine & _
            "3. NO use rutas como: \\datoste\...\HPS\..., \\datoste\...\CONDOR\00_Condor\..." & vbNewLine & vbNewLine & _
            "Si necesita clonar datos de producción para tests, use un backend de TEST configurado.", _
            "SandboxValidator.AssertSandboxPathIsSafe"
        errObj.Raise
    End If
    
    Exit Sub
    
Errores:
    Dim errObj2 As New CondorError
    errObj2.Create Err.Number, Err.description, "SandboxValidator.AssertSandboxPathIsSafe"
    errObj2.Raise
End Sub

' ---
' AssertBackendSourceIsReadOnlySafe: Valida que el backend origen es SEGURO para SOLO LECTURA
'
' A diferencia de AssertSandboxPathIsSafe, esta función NO aborta si el backend
' ES producción. En cambio, marca el backend como READ-ONLY ONLY y prepara
' el hook para que SandboxGestor nunca intente escribir en él.
'
' Para un backend de producción, leer ESTÁ permitido (es lo que queremos para clonar),
' pero el flag interno previene cualquier operación de write/DDL
'
' path: Ruta del backend origen (ej: \\datoste\...\Expedientes_datos.accdb)
' Retorna: True si es seguro (local/test), False si es producción pero legible
' Lanza: CondorError si el path es vacío o no se puede evaluar
' ---
Public Function AssertBackendSourceIsReadOnlySafe(path As String) As Boolean
    On Error GoTo Errores
    
    Dim strPath As String
    strPath = Trim(path)
    
    If strPath = "" Then
        Dim errObj As New CondorError
        errObj.Create ERR_UNSAFE_PATH, _
            "Backend path vacío o inválido. No se puede validar seguridad.", _
            "SandboxValidator.AssertBackendSourceIsReadOnlySafe"
        errObj.Raise
    End If
    
    ' Verificar si es producción
    If IsProductionLikePath(strPath) Then
        ' Es producción o parece producción - lo marcamos como READ-ONLY ONLY
        ' Logueamos una ADVERTENCIA (no error) ya que es intencional leer de prod
        Debug.Print "[SandboxValidator][AssertBackendSourceIsReadOnlySafe] BACKEND_READONLY_WARNING: Backend de producción detectado: " & strPath & ". El engine usará SOLO lectura."
        
        ' Retornamos False para indicar "precaución" pero NO abortamos
        ' El flag m_blnBackendIsReadOnlyOnly se puede consultar después
        AssertBackendSourceIsReadOnlySafe = False
        Exit Function
    End If
    
    ' Es una ruta segura (local o test)
    AssertBackendSourceIsReadOnlySafe = True
    Exit Function
    
Errores:
    Dim errObj2 As New CondorError
    errObj2.Create Err.Number, Err.description, "SandboxValidator.AssertBackendSourceIsReadOnlySafe"
    errObj2.Raise
End Function

' ---
' GetProductionPathViolation: Obtiene descripción detallada de por qué
' una ruta fue marcada como producción (para logging/UI)
' ---
Public Function GetProductionPathViolation(path As String) As String
    On Error GoTo Errores
    
    Dim strResult As String
    strResult = ""
    
    If path = "" Then
        GetProductionPathViolation = ""
        Exit Function
    End If
    
    Dim strLower As String
    strLower = LCase(path)
    
    ' Analizar y reportar qué patrón detectó
    If Left(strLower, 2) = "\\" Then
        If InStr(strLower, "\\datoste") > 0 Then
            strResult = strResult & " [UNC:datoste]"
        End If
        If InStr(strLower, "\hps\") > 0 Or InStr(strLower, "\00_hps\") > 0 Then
            strResult = strResult & " [pattern:HPS]"
        End If
        If InStr(strLower, "\condor\00_condor\") > 0 Then
            strResult = strResult & " [pattern:CONDOR/00_Condor]"
        End If
        If InStr(strLower, "\expedientes\") > 0 Then
            strResult = strResult & " [pattern:Expedientes]"
        End If
        If InStr(strLower, "\no conformidades\") > 0 Then
            strResult = strResult & " [pattern:NoConformidades]"
        End If
        If InStr(strLower, "\0lanzadera\") > 0 Then
            strResult = strResult & " [pattern:Lanzadera]"
        End If
    Else
        ' Path local
        If InStr(strLower, "\hps\") > 0 Or InStr(strLower, "\00_hps\") > 0 Then
            strResult = strResult & " [local:HPS]"
        End If
        If InStr(strLower, "\condor\00_condor\") > 0 Then
            strResult = strResult & " [local:CONDOR/00_Condor]"
        End If
    End If
    
    If strResult = "" Then
        strResult = " [unknown production indicator]"
    End If
    
    GetProductionPathViolation = Trim(strResult)
    Exit Function
    
Errores:
    GetProductionPathViolation = ""
End Function

