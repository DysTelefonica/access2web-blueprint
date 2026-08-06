Attribute VB_Name = "Variables Globales"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Variables Globales.bas (VERSIÓN LIMPIA SIN FUNCIONES DUPLICADAS)
' ==========================================================================

' --- NOTA: Todas las declaraciones de API de Windows se han movido a "FUNCIONES UTILES.bas" ---


Public Const PROJECT_NAME As String = "CONDOR"

Public m_TituloFormulario As String


Public fso As New FileSystemObject

Public Const msoFileDialogFilePicker As Long = 3
Public Const msoFileDialogFolderPicker As Long = 4
Public Const msoFileDialogOpen As Long = 1
Public Const msoFileDialogSaveAs As Long = 2
Public Const SubRedOficina As String = "10.14.7"

' --- OBJETOS GLOBALES SINGLETON ---
Public m_ObjEntorno As Entorno
Public m_ObjUsuarioReal As usuario
Public m_ObjUsuarioActivo As usuario

' --- GESTIÓN DE ESPACIO DE TRABAJO (SOLUCIÓN BLOQUEOS) ---
Public g_wsCondor As DAO.Workspace

' Gestores de Conexiones a Bases de Datos
Public g_dbCondor As DAO.Database
Public g_dbLanzadera As DAO.Database
Public g_dbCorreos As DAO.Database
Public g_dbExpedientes As DAO.Database
Public g_dbNoConformidades As DAO.Database

' LEGACY — transitional only. Used by test harness to override getdb() routing.
' After test-sandbox-via-backend-switching, this should be removed entirely.
' NOT reset in ResetGlobals() — intentional per design.
Public g_dbTestInstance As DAO.Database

' --- BACKEND CONFIG (module-level vars shared across modules) ---
Public m_ActiveBackendURL As String
Public m_PasswordBackend As String
Public m_dbCached As DAO.Database

' --- TEST ROUTING STATE ---
' False by default. Test helpers set this to True only after validating a
' sandbox/local backend. Do not reset in ResetGlobals(); it is session-scoped
' test lifecycle state, not production backend configuration.
Public m_TestingMode As Boolean
Public m_BackendSandboxURL As String

' --- TEST SANDBOX PASSWORD (PR2, Spec-005/008) ---
' Canonical alias for the asserted sandbox password resolved by
' TestHelper.AssertSandboxBackend. Kept in lockstep with m_PasswordBackend
' so consumers can read either; m_PasswordBackend remains the production
' routing var (read from TbConfiguracionBackends.PasswordBackend) and
' m_BackendSandboxPassword is the test-side name. ResetTestSession clears
' only the test caches and never the production cache.
Public m_BackendSandboxPassword As String

' --- VARIABLES GLOBALES DE ESTADO ---
Public g_blnImpersonando As Boolean
Public rolUsuario As rol
Public rolUsuarioReal As rol ' <-- AÑADIR ESTA LÍNEA
Public IDAplicacion As String

Public m_URLRutaAplicacionesLocal As String
Public m_URLRutaAplicacionesRemotas As String
Public m_URLRutaAplicacionLocal As String
Public m_URLRutaAplicacionRemota As String
Public g_objLastError As CondorError ' Variable global para transportar el objeto de error
' --- VARIABLES GLOBALES TEMPORALES (USO LIMITADO) ---
Public t1 As Single
Public t2 As Single

Public Sub EVE()
    Dim userServ As New UsuarioServicio
    Dim errObj As New CondorError
    Dim m_Command As String
    Dim t1 As Single
    Dim t2 As Single
    
    On Error GoTo Errores
    t1 = Timer
    ' --- ETAPA 0: INICIALIZACIÓN DEL WORKSPACE (CRÍTICO) ---
    ' Creamos un espacio de trabajo propio para aislar nuestras transacciones de la UI de Access
    Avance "Inicializando motor de datos..."
    Set g_wsCondor = DBEngine.CreateWorkspace("CondorWS", "admin", "")

    ' --- PR2 (Spec-005/008): test-mode guard ---------------------------------
    ' When m_TestingMode is already active, EVE is being driven by the test
    ' harness (TestHelper.BeginTestSession) which has already validated and
    ' cached m_BackendSandboxURL / m_BackendSandboxPassword / m_PasswordBackend.
    ' Skipping ResetGlobals + LeeConfiguracionLocal here preserves the test
    ' state and prevents LeeConfiguracionLocal from overwriting the sandbox
    ' routing with the production BackendActivo=PROD row. Routing through
    ' m_BackendSandboxURL is delegated to CheckAndReconnect's test-mode branch.
    Dim cfgError As String
    If m_TestingMode Then
        cfgError = ""
        Avance "Modo testing activo: saltando ResetGlobals/LeeConfiguracionLocal..."
        ' --- PR2 follow-up: in test mode EVE is a "soft init" that sets up the
        ' workspace and returns. Sections 1.1-1.3 (backend connectivity probe,
        ' user identification/permission check) and the Entorno validation block
        ' expect production state (a real user row, an INI/config file, etc.)
        ' that the sandbox does not have. BeginTestSession has already validated
        ' the sandbox and cached the URL/password, so this early-return keeps
        ' the test-mode path independent of the production User/Entorno stack.
        Avance "Modo testing activo: skip validación 1.1-1.3 y carga de Entorno..."
        Exit Sub
    Else
        ' --- Backend config reset and read (MUST be first after workspace init) ---
        Call ResetGlobals
        Call LeeConfiguracionLocal(cfgError)
        If cfgError <> "" Then
            Err.Raise 513, "EVE.LeeConfiguracionLocal", cfgError
        End If
    End If
    
    ' --- ETAPA 1: VALIDACIÓN DE DEPENDENCIAS CRÍTICAS (Pre-arranque) ---
    Avance "Validando dependencias críticas..."
    
    ' Inicialización básica de variables
    Application.TempVars("EnDesarrollo") = "Sí"
    Application.TempVars("DatosEnLocal") = "No"
    Application.TempVars("EnPruebas") = "No"
    
    ' --- NUEVO FLAG ---
    ' "Sí" = Usa archivos temporales (Más estable, evita bloqueo)
    ' "No" = Usa inyección en memoria (Método antiguo)
    Application.TempVars("UsarRenderizadoFichero") = "Sí"
    
    ' SUSTITUCIÓN DE NÚMERO MÁGICO "23"
    IDAplicacion = CONST_ID_APLICACION
    
    m_URLRutaAplicacionesRemotas = "\\datoste\aplicaciones_dys\Aplicaciones PpD\"
    m_URLRutaAplicacionRemota = m_URLRutaAplicacionesRemotas & "CONDOR\"
    m_URLRutaAplicacionesLocal = "C:\Users\adm1\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\"
    m_URLRutaAplicacionLocal = m_URLRutaAplicacionesLocal & "CONDOR\"
    
    ' 1.1: Validar conectividad con el backend configurado
    On Error Resume Next
    Dim dbTest As DAO.Database
    Set dbTest = getdb()
    If Err.Number <> 0 Then
        Err.Raise 513, "EVE.ValidacionCritica", "No se pudo establecer conexión con el backend configurado (" & Nz(Application.TempVars("BackendActivo"), "PROD") & ")."
    End If
   
    On Error GoTo Errores
    
    ' 1.2: Obtener el usuario
    Avance "Identificando usuario..."
    m_Command = Nz(VBA.Command, "")
    If m_Command <> "" Then
        Set m_ObjUsuarioReal = userServ.getUsuarioConPermisos(p_Correo:=m_Command)
    Else
        Set m_ObjUsuarioReal = userServ.getUsuarioConectadoConPermisos()
    End If
    
    ' 1.3: Validar que el usuario existe
    If m_ObjUsuarioReal Is Nothing Then
        Err.Raise 513, "EVE.ValidacionCritica", "Su usuario de red no está registrado en la aplicación o no tiene permisos."
    End If
    
    Set m_ObjUsuarioActivo = m_ObjUsuarioReal
    g_blnImpersonando = False
    rolUsuario = m_ObjUsuarioReal.rol
    rolUsuarioReal = m_ObjUsuarioReal.rol
    
    ' --- ETAPA 2: CARGA Y VALIDACIÓN DEL ENTORNO ---
    Avance "Cargando configuración del entorno..."
    Set m_ObjEntorno = New Entorno
    
    Dim colErrores As Object: Set colErrores = CreateObject("Scripting.Dictionary")
    Dim m_NombreCampo As Variant
    Dim m_valor As String
    
    For Each m_NombreCampo In m_ObjEntorno.ColItems.Keys
        Avance "Verificando: " & m_NombreCampo
        On Error Resume Next
        m_valor = m_ObjEntorno.getPropiedad(CStr(m_NombreCampo))
        
        If Err.Number <> 0 Then
            colErrores.Add CStr(m_NombreCampo), Err.description
            Err.Clear
        End If
        On Error GoTo Errores
    Next
    
    If colErrores.count > 0 Then
        Dim msgFinal As String
        msgFinal = "CONDOR no pudo iniciarse debido a errores de configuración."
        Err.Raise 513, "EVE.ValidacionEntorno", msgFinal
    End If
    t2 = Timer
    Debug.Print t2 - t1; " segundos"
    
    Exit Sub

Errores:
    Set errObj = New CondorError
    errObj.Create Err.Number, Err.description, "Variables Globales.EVE"
    errObj.Raise
End Sub

Public Sub ResetGlobals()
    ' TempVars — clear all for fresh start
    Application.TempVars.RemoveAll

    ' Backend config — reset to empty so next getdb() re-reads from table
    m_ActiveBackendURL = ""
    m_PasswordBackend = ""

    ' Connections — clear caches
    Set m_dbCached = Nothing
    Set g_dbCondor = Nothing

    ' Workspace reset (critical for fresh start)
    Set g_wsCondor = Nothing
    
    ' m_Obj* globals reset (required by spec)
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioReal = Nothing
    Set m_ObjUsuarioActivo = Nothing

    ' Legacy connection globals (unused after migration, declarations kept)
    ' g_dbLanzadera, g_dbCorreos, g_dbExpedientes, g_dbNoConformidades
End Sub

Public Function LeeConfiguracionLocal( _
                Optional ByRef p_Error As String _
                ) As String
    Dim rcdCfg As DAO.Recordset
    Dim m_SQL As String
    Dim m_BackendActivo As String
    Dim m_URL As String
    Dim m_Pwd As String
    Dim m_RutaDirectorioAplicacion As String

    On Error GoTo Errores

    m_SQL = "SELECT TOP 1 * FROM TbConfiguracionBackends;"
    Set rcdCfg = CurrentDb.OpenRecordset(m_SQL)

    If rcdCfg.EOF Then
        rcdCfg.Close: Set rcdCfg = Nothing
        p_Error = "No se encontro configuracion de backend habilitada"
        Exit Function
    End If

    rcdCfg.MoveFirst

    ' Backend activo
    m_BackendActivo = Nz(rcdCfg!BackendActivo, "PROD")
    Application.TempVars("BackendActivo") = m_BackendActivo
    Application.TempVars("BackendProduccion") = Nz(rcdCfg!BackendProduccion, "")
    Application.TempVars("BackendSandbox") = Nz(rcdCfg!BackendSandbox, "")
    Application.TempVars("BackendTest") = Nz(rcdCfg!BackendTest, "")

    ' Password
    m_Pwd = Nz(rcdCfg!PasswordBackend, "")
    Application.TempVars("PasswordBackend") = m_Pwd

    ' Ruta de recursos segun entorno (EnPruebas)
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_RutaDirectorioAplicacion = Nz(rcdCfg!RutaDirectorioAplicacion_LOCAL, "")
    Else
        m_RutaDirectorioAplicacion = Nz(rcdCfg!RutaDirectorioAplicacion_PROD, "")
    End If
    Application.TempVars("RutaDirectorioAplicacion") = m_RutaDirectorioAplicacion

    ' URL del backend segun entorno activo
    Select Case m_BackendActivo
        Case "PROD":    m_URL = Nz(rcdCfg!BackendProduccion, "")
        Case "SANDBOX": m_URL = Nz(rcdCfg!BackendSandbox, "")
        Case "TEST":    m_URL = Nz(rcdCfg!BackendTest, "")
        Case Else:      m_URL = Nz(rcdCfg!BackendProduccion, "")
    End Select

    ' Guardar en variables de modulo (para getdb)
    m_ActiveBackendURL = m_URL
    m_PasswordBackend = m_Pwd

    LeeConfiguracionLocal = m_URL

    rcdCfg.Close: Set rcdCfg = Nothing
    Exit Function

Errores:
    If Not rcdCfg Is Nothing Then rcdCfg.Close: Set rcdCfg = Nothing
    If Err.Number <> 1000 Then
        p_Error = "LeeConfiguracionLocal: " & Err.description
    End If
End Function

Public Function LeerIni(key As String, Default As Variant) As String
    Dim bufer As String * 256, Len_Value As Long
    Len_Value = GetPrivateProfileString(fso.GetBaseName(CurrentDb().name), key, Default, bufer, Len(bufer), m_ObjEntorno.URLAchivoIni)
    LeerIni = Left$(bufer, CLng(Len_Value))
End Function

 
' ==========================================================================
' MÓDULO: Variables Globales.bas
' ACCIÓN: Añadir procedimiento para refrescar contadores globales.
' ==========================================================================
Public Sub RefrescarTablerosPrincipales()
    ' RESPONSABILIDAD: Actualizar los contadores de tareas en los menús principales
    '                  si están abiertos, independientemente del rol.
    On Error Resume Next
    
    If FormularioAbierto("frm0Ppal") Then
        Forms("frm0Ppal").ActualizarContadorTareas
    End If
    
    If FormularioAbierto("frm0PpalTecnico") Then
        Forms("frm0PpalTecnico").ActualizarContadorTareas
    End If
End Sub


