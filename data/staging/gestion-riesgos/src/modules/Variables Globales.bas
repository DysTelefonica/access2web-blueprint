Attribute VB_Name = "Variables Globales"
Option Compare Database
Option Explicit

#If Win64 = 1 Then
    
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
            ByVal lpApplicationName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As String, _
            ByVal nSize As Long, _
            ByVal lpFileName As String) As Long
    
    Public Declare PtrSafe Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    Public Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" ( _
        ByVal hProcess As Long, lpExitCode As Long) As Long
    Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As Long)
    Public Declare PtrSafe Function GetIpAddrTable Lib "Iphlpapi" ( _
            pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
#Else
    
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
            ByVal lpApplicationName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As String, _
            ByVal nSize As Long, _
            ByVal lpFileName As String) As Long
    Public Declare Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    Public Declare Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Public Declare Function GetExitCodeProcess Lib "kernel32" ( _
        ByVal hProcess As Long, lpExitCode As Long) As Long
    Public Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    Public Declare  Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As Long)
    Public Declare  Function GetIpAddrTable Lib "Iphlpapi" ( _
            pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
#End If



Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STATUS_PENDING = &H103&


Public pregunta As Long
Public lbl As Label
Public fso As New FileSystemObject
Public Const SubRedOficina As String = "10.14.7"
Public m_ObjEntorno As Entorno

Public m_ObjProyectoActivo As Proyecto
Public m_ObjProyectoAlInicio As Proyecto


Public m_ObjEdicionActiva As Edicion
Public m_ObjRiesgoActivo As riesgo
Public m_EstadoRiesgoActivo As EnumRiesgoEstado
Public m_EsAlta As EnumSiNo
Public m_ObjRiesgoAlInicio As riesgo
Public blnEdicionActiva As Boolean
'Public blnPermitidoEditar As Boolean
Public m_ObjRiesgoExtActivo As RiesgoExterno
Public m_ObjPMActivo As PM
Public m_ObjPCActivo As PC
Public m_ObjPMAccionActiva As PMAccion
Public m_ObjPCAccionActiva As PCAccion
Public m_ObjNCActiva As NC
Public m_ObjSuministradorActivo As Suministrador
Public m_ObjRiesgoMaterializadoActivo As RiesgoMaterializacion
Public m_ObjRiesgoBibliotecaActivo As RiesgoBiblioteca



Public m_ObjTareasCalidad As TareasCalidad
Public m_ObjTareasTecnico As TareasTecnico
Public wks As DAO.Workspace
Private db As DAO.Database
Private db1 As DAO.Database
Private m_CachedDB As DAO.Database
Public m_IDAplicacion As String
Public m_ObjUsuarioConectadoInicialmente As Usuario
Public m_ObjUsuarioConectado As Usuario
Public EsAdministrador As EnumSiNo
Public EsAdministradorConectadoInicialmente As EnumSiNo
Public EsCalidad As EnumSiNo
Public EsTecnico As EnumSiNo

' Cache global de entidades (lazy init — se crean en GetCachedX, no aqui)
Private m_DicRiesgos As Scripting.Dictionary
Private m_DicEdiciones As Scripting.Dictionary
Private m_DicProyectos As Scripting.Dictionary
Private m_DicPMs As Scripting.Dictionary
Private m_DicPCs As Scripting.Dictionary

Public m_DBOpen As Boolean
Public m_ActiveBackendURL As String   ' URL del backend activo — para tests y acceso directo
Public m_PasswordBackend As String   ' password del backend — para tests y acceso directa
Public m_BackendSandboxURL As String  ' sandbox URL para tests — siempre disponible, independientemente de BackendActivo
Public m_TestingMode As Boolean       ' True solo durante tests: getdb usa exclusivamente m_BackendSandboxURL
Public m_EnPruebas As Boolean
Public m_RutaDirApp As String
Public m_RutaDirApp_PROD As String
Public m_RutaDirApp_LOCAL As String
Public m_BackendActivo As String  ' "PROD" or "LOCAL"/"SANDBOX"
Public IDAplicacion As String  ' legacy accessor — usar m_IDAplicacion en código nuevo

' TempVars config vars (mirror de TbConfiguracionBackends)
Public m_CadenaJerarquicaModelo As String
Public m_JPMesesAvisoEntreEdiciones As Long
Public m_JPDiasPreviosParaElAviso As Long
Public m_CalDiaInicialMesAviso As Long
Public m_Publicabilidad_Usar_Cache As String

Public t1 As Single
Public t2 As Single
Public varItem As Variant
Public m_EnOficina As EnumSiNo

Public m_ObjUsuarioParaTareas As Usuario

Public blnPermitidoEscribir As Boolean
Public m_ObjAnexoEvicenciaUTE As Anexo
Public m_URLInforme As String
Public m_URLHTMLActivo As String
Public m_URLRutaAplicacionesLocal As String
Public m_URLRutaAplicacionesRemotas As String
Public m_URLRutaAplicacionLocal As String
Public m_URLRutaAplicacionRemota As String
Public m_ListaUsuarios As String
Public m_ObjUltimoProyecto As UltimoProyecto

' --- GetCachedRiesgo ---
Public Function GetCachedRiesgo(ByVal p_IDRiesgo As String, _
                                Optional ByRef p_Error As String, _
                                Optional ByRef db As DAO.Database = Nothing) As riesgo
    On Error GoTo errores
    If p_IDRiesgo = "" Then Exit Function
    
    If m_DicRiesgos Is Nothing Then
        Set m_DicRiesgos = New Scripting.Dictionary
        m_DicRiesgos.CompareMode = TextCompare
    End If
    
    If m_DicRiesgos.Exists(p_IDRiesgo) Then
        Set GetCachedRiesgo = m_DicRiesgos(p_IDRiesgo)
    Else
        Dim obj As riesgo
        Set obj = Constructor.getRiesgo(p_IDRiesgo, , , p_Error, db)
        If Not obj Is Nothing Then
            m_DicRiesgos.Add p_IDRiesgo, obj
            Set GetCachedRiesgo = obj
        End If
    End If
    Exit Function
errores:
    p_Error = "GetCachedRiesgo: " & Err.Number & " - " & Err.Description
End Function

Public Sub InvalidarCacheRiesgo(ByVal p_IDRiesgo As String, Optional ByRef p_Error As String)
    Dim m_ErrorLocal As String
    On Error GoTo errores

    p_Error = ""
    If Not m_DicRiesgos Is Nothing Then
        If m_DicRiesgos.Exists(p_IDRiesgo) Then
            m_DicRiesgos.Remove p_IDRiesgo
        End If
    End If
    m_ErrorLocal = ""
    CachePublicabilidad_InvalidarRiesgo p_IDRiesgo:=p_IDRiesgo, p_Error:=m_ErrorLocal
    If m_ErrorLocal <> "" Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        p_Error = "InvalidarCacheRiesgo: " & Err.Number & " - " & Err.Description
    End If
End Sub

' --- GetCachedRiesgoFresh ---
' Garantiza instancia DISTINTA de cualquier instancia previamente cacheada
' para el mismo p_IDRiesgo. Contrato: la instancia devuelta tiene una
' referencia de memoria diferente de cualquier GetCachedRiesgo previo.
'
' Implementado como: InvalidarCacheRiesgo + GetCachedRiesgo. Sin este
' patron, dos llamadas consecutivas a GetCachedRiesgo devuelven la MISMA
' instancia (cache hit sobre la misma referencia), rompiendo el baseline
' de diff en RiesgoChangeDetector.HasChanges (ver Form_FormRiesgo.cls:266
' bug latente).
'
' Validaciones:
'   - p_IDRiesgo vacio o whitespace: Nothing + p_Error mencionando "vacio"
'   - p_IDRiesgo no encontrado en DB: Nothing + p_Error propagado de GetCachedRiesgo
'   - Cualquier fallo en InvalidarCacheRiesgo: Nothing + p_Error propagado
Public Function GetCachedRiesgoFresh(ByVal p_IDRiesgo As String, _
                                     Optional ByRef p_Error As String, _
                                     Optional ByRef db As DAO.Database = Nothing) As riesgo
    On Error GoTo errores
    p_Error = ""

    ' Validacion explicita: IDs vacios o whitespace son invalidos
    If Len(Trim$(p_IDRiesgo)) = 0 Then
        p_Error = "GetCachedRiesgoFresh: p_IDRiesgo esta vacio"
        Exit Function
    End If

    ' 1. Invalidar cache para garantizar instancia fresca en el siguiente fetch
    InvalidarCacheRiesgo p_IDRiesgo, p_Error
    If p_Error <> "" Then
        Exit Function
    End If

    ' 2. GetCachedRiesgo hace cache miss -> Constructor.getRiesgo -> New riesgo
    '    Si el ID no existe en DB, GetCachedRiesgo retorna Nothing y p_Error
    '    queda propagado desde Constructor.getRiesgo.
    Set GetCachedRiesgoFresh = GetCachedRiesgo(p_IDRiesgo, p_Error, db)

    ' FIX W1-S4: si GetCachedRiesgo retorno Nothing sin popular p_Error
    ' (caso Constructor.getRiesgo EOF silencioso), poblar p_Error explicitamente
    ' para honrar convencion p_Error ByRef de Telefonica D&S.
    If GetCachedRiesgoFresh Is Nothing And p_Error = "" Then
        p_Error = "GetCachedRiesgoFresh: ID '" & p_IDRiesgo & "' no encontrado en TbRiesgos"
    End If

    Exit Function
errores:
    p_Error = "GetCachedRiesgoFresh: " & Err.Number & " - " & Err.Description
End Function

' --- CacheRiesgoContainsKey (test-only accessor) ---
' Returns True iff m_DicRiesgos currently contains p_IDRiesgo.
' Test-only: production code should NEVER use this to inspect cache state
' (bypassing GetCachedRiesgo defeats the cache contract).
' Used by Test_RiesgoEstadoGateHelper W2 atoms to assert cache population
' after gate helper lookups (Constructor.getRiesgo vs GetCachedRiesgo).
Public Function CacheRiesgoContainsKey(ByVal p_IDRiesgo As String) As Boolean
    On Error GoTo errores
    If m_DicRiesgos Is Nothing Then
        CacheRiesgoContainsKey = False
    Else
        CacheRiesgoContainsKey = m_DicRiesgos.Exists(p_IDRiesgo)
    End If
    Exit Function
errores:
    CacheRiesgoContainsKey = False
End Function

' --- GetCachedEdicion ---
Public Function GetCachedEdicion(ByVal p_IDEdicion As String, _
                                  Optional ByRef p_Error As String) As Edicion
    On Error GoTo errores
If p_IDEdicion = "" Then Exit Function

    If m_DicEdiciones Is Nothing Then
        Set m_DicEdiciones = New Scripting.Dictionary
        m_DicEdiciones.CompareMode = TextCompare
    End If

    If m_DicEdiciones.Exists(p_IDEdicion) Then
        Set GetCachedEdicion = m_DicEdiciones(p_IDEdicion)
    Else
        Dim obj As Edicion
        Set obj = Constructor.getEdicion(p_IDEdicion, p_Error)
        If Not obj Is Nothing Then
            m_DicEdiciones.Add p_IDEdicion, obj
            Set GetCachedEdicion = obj
        End If
    End If
    Exit Function
errores:
    p_Error = "GetCachedEdicion: " & Err.Number & " - " & Err.Description
End Function

Public Sub InvalidarCacheEdicion(ByVal p_IDEdicion As String, Optional ByRef p_Error As String)
    Dim m_ErrorLocal As String
    On Error GoTo errores

    p_Error = ""
    If Not m_DicEdiciones Is Nothing Then
        If m_DicEdiciones.Exists(p_IDEdicion) Then
            m_DicEdiciones.Remove p_IDEdicion
        End If
    End If
    m_ErrorLocal = ""
    CachePublicabilidad_InvalidarEdicion p_IDEdicion:=p_IDEdicion, p_Error:=m_ErrorLocal
    If m_ErrorLocal <> "" Then
        p_Error = m_ErrorLocal
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        p_Error = "InvalidarCacheEdicion: " & Err.Number & " - " & Err.Description
    End If
End Sub

Public Sub InvalidarPublicabilidadPorCambioEvidenciaEdicion(ByVal p_IDEdicion As String, Optional ByRef p_Error As String)
    On Error GoTo errores

    p_Error = ""
    If p_IDEdicion = "" Then
        Exit Sub
    End If

    InvalidarCacheEdicion p_IDEdicion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        p_Error = "InvalidarPublicabilidadPorCambioEvidenciaEdicion: " & Err.Number & " - " & Err.Description
    End If
End Sub

' --- GetCachedProyecto ---
Public Function GetCachedProyecto(ByVal p_IDProyecto As String, _
                                   Optional ByRef p_Error As String) As Proyecto
    On Error GoTo errores
    If p_IDProyecto = "" Then Exit Function
    
    If m_DicProyectos Is Nothing Then
        Set m_DicProyectos = New Scripting.Dictionary
        m_DicProyectos.CompareMode = TextCompare
    End If
    
    If m_DicProyectos.Exists(p_IDProyecto) Then
        Set GetCachedProyecto = m_DicProyectos(p_IDProyecto)
    Else
        Dim obj As Proyecto
        Set obj = Constructor.getProyecto(p_IDProyecto, p_Error)
        If Not obj Is Nothing Then
            m_DicProyectos.Add p_IDProyecto, obj
            Set GetCachedProyecto = obj
        End If
    End If
    Exit Function
errores:
    p_Error = "GetCachedProyecto: " & Err.Number & " - " & Err.Description
End Function

Public Sub InvalidarCacheProyecto(ByVal p_IDProyecto As String)
    If Not m_DicProyectos Is Nothing Then
        If m_DicProyectos.Exists(p_IDProyecto) Then
            m_DicProyectos.Remove p_IDProyecto
        End If
    End If
End Sub

' --- GetCachedPM ---
Public Function GetCachedPM(ByVal p_IDPM As String, _
                             Optional ByRef p_Error As String) As PM
    On Error GoTo errores
    If p_IDPM = "" Then Exit Function
    
    If m_DicPMs Is Nothing Then
        Set m_DicPMs = New Scripting.Dictionary
        m_DicPMs.CompareMode = TextCompare
    End If
    
    If m_DicPMs.Exists(p_IDPM) Then
        Set GetCachedPM = m_DicPMs(p_IDPM)
    Else
        Dim obj As PM
        Set obj = Constructor.getPM(p_IDPM, p_Error)
        If Not obj Is Nothing Then
            m_DicPMs.Add p_IDPM, obj
            Set GetCachedPM = obj
        End If
    End If
    Exit Function
errores:
    p_Error = "GetCachedPM: " & Err.Number & " - " & Err.Description
End Function

Public Sub InvalidarCachePM(ByVal p_IDPM As String)
    If Not m_DicPMs Is Nothing Then
        If m_DicPMs.Exists(p_IDPM) Then
            m_DicPMs.Remove p_IDPM
        End If
    End If
End Sub

' --- GetCachedPC ---
Public Function GetCachedPC(ByVal p_IDPC As String, _
                             Optional ByRef p_Error As String) As PC
    On Error GoTo errores
    If p_IDPC = "" Then Exit Function
    
    If m_DicPCs Is Nothing Then
        Set m_DicPCs = New Scripting.Dictionary
        m_DicPCs.CompareMode = TextCompare
    End If
    
    If m_DicPCs.Exists(p_IDPC) Then
        Set GetCachedPC = m_DicPCs(p_IDPC)
    Else
        Dim obj As PC
        Set obj = Constructor.getPC(p_IDPC, p_Error)
        If Not obj Is Nothing Then
            m_DicPCs.Add p_IDPC, obj
            Set GetCachedPC = obj
        End If
    End If
    Exit Function
errores:
    p_Error = "GetCachedPC: " & Err.Number & " - " & Err.Description
End Function

Public Sub InvalidarCachePC(ByVal p_IDPC As String)
    If Not m_DicPCs Is Nothing Then
        If m_DicPCs.Exists(p_IDPC) Then
            m_DicPCs.Remove p_IDPC
        End If
    End If
End Sub

' --- GetCachedRiesgo ---


Public Function getNombreUsuarioConectado(Optional ByRef p_Error As String) As String
    
    Dim m_UsuarioMaquina As Usuario
    
    On Error GoTo errores
    
    If Not m_ObjUsuarioConectado Is Nothing Then
        getNombreUsuarioConectado = m_ObjUsuarioConectado.Nombre
        Exit Function
    End If
    
    Set m_UsuarioMaquina = Constructor.getUsuarioConectadoPorMaquina(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_UsuarioMaquina Is Nothing Then
        getNombreUsuarioConectado = "Desconocido"
        Exit Function
    End If
    getNombreUsuarioConectado = m_UsuarioMaquina.Nombre
    Exit Function
errores:
    getNombreUsuarioConectado = "Desconocido"
End Function
Public Function ReiniciarLasVariables(Optional ByRef p_Error As String) As String
    
    
    
    On Error GoTo errores
   
    Set m_ObjProyectoActivo = Nothing
    Set m_ObjProyectoAlInicio = Nothing
    Set m_ObjEdicionActiva = Nothing
    Set m_ObjRiesgoActivo = Nothing
    m_EsAlta = Empty
    Set m_ObjRiesgoAlInicio = Nothing
    blnEdicionActiva = False
    
    
    
    Set m_ObjRiesgoExtActivo = Nothing
    Set m_ObjPMActivo = Nothing
    Set m_ObjPMAccionActiva = Nothing
    Set m_ObjPCActivo = Nothing
    Set m_ObjPCAccionActiva = Nothing
    
    Set m_ObjNCActiva = Nothing
    Set m_ObjSuministradorActivo = Nothing
    Set m_ObjRiesgoMaterializadoActivo = Nothing
    Set m_ObjRiesgoBibliotecaActivo = Nothing
    
    Set m_ObjTareasCalidad = Nothing
    Set m_ObjTareasTecnico = Nothing
    Set m_ObjUsuarioConectado = Nothing
    EsAdministrador = Empty
    EsCalidad = Empty
    EsTecnico = Empty
    
    Set m_ObjUsuarioParaTareas = Nothing
    blnPermitidoEscribir = False
    Set m_ObjAnexoEvicenciaUTE = Nothing
    m_URLInforme = ""
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo ReiniciarLasVariables ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function


Public Function EVE( _
                    Optional ByRef p_CorreoUsuario As String, _
                    Optional ByRef p_Error As String _
                    ) As String

    Dim m_NombreCarpeta As String
    Dim m_NombreCampo As Variant
    Dim m_valor As String
    Dim m_Objeto As Object
    Dim ti As Single
    Dim tf As Single
    Dim m_UsuarioLogeadoEnOrdenador As String
    
    Dim m_TipoCampo As String
    Dim m_ValorCampo As String
   
    Dim m_ValorCampoTruncado As String
    Dim m_Command As String 'se obtienen cuando se abre la base de datos con parámteros
    Dim m_ComandoResultante As String
    Dim m_Linea As String
    Dim m_ClaveValor As String
    Dim objNetwork As Object
    
    Dim intNumeroErrores As Integer
    Dim m_CadenaCamposConError As String
    On Error GoTo errores
    
    ti = Timer
    
    ' 1. FIRST: Reset globals and clear TempVars — start clean
    Application.TempVars.RemoveAll
    ResetGlobals p_Error
    If p_Error <> "" Then
        p_Error = "EVE: Error en ResetGlobals: " & p_Error
        Err.Raise 1000
    End If
    
    ' 2. THEN: Load backend configuration from table
    LeeConfiguracionLocal p_Error
    If p_Error <> "" Then
        MsgBox "ERROR CRÍTICO: " & p_Error, vbCritical, "EVE"
        Exit Function
    End If
    
    ' 3. THEN: Initialize Entorno (needs m_RutaDirApp from config)
    Set m_ObjEntorno = New Entorno
    ' Validar TODOS los recursos críticos ycollect errores consolidado
    Dim erroresRecursos As String
    erroresRecursos = m_ObjEntorno.ValidarRecursos()
    If erroresRecursos <> "" Then
        p_Error = "RECURSOS INACCESIBLES:" & vbCrLf & erroresRecursos
        Err.Raise 1000
    End If
    ' m_URLRutaAplicacionRemota y m_URLRutaAplicacionLocal son el directorio de aplicación completo
    m_URLRutaAplicacionRemota = m_RutaDirApp
    m_URLRutaAplicacionLocal = m_RutaDirApp_LOCAL
    ' m_URLRutaAplicacionesRemotas y m_URLRutaAplicacionesLocal ya fueron derivados
    ' de la tabla en LeeConfiguracionLocal — no sobreescribir aquí
    m_Command = Nz(VBA.Command, "")
   ' m_Command = "beatriz.novalgutierrez@telefonica.com"
   ' m_Command = "marta.garridovaamonde@telefonica.com"
    'm_Command = "rosamaria.fuentesherrero@telefonica.com"
    'm_Command = "felix.sanchezpimentel@telefonica.com"
    'm_Command = "sergio.garciamontalvo@telefonica.com"
    'm_Command = "juan.jerezgarcia@telefonica.com"
    'm_Command = "javier.amousanos@telefonica.com"
    'm_Command = "jose.perezdionisio@telefonica.com"
    'm_Command = "carlos.alonsocarmona@telefonica.com"
    'm_Command = "juliobenedicto.vicariomancho@telefonica.com"
    'm_Command = "mario.martinabad@telefonica.com"
    'm_Command = "anamaria.rubiocanales@telefonica.com"
    'm_Command = "natalia.casangarcia@telefonica.com"
    'm_Command = "fernando.lazarodiaz@telefonica.com"
    t1 = Timer
    If p_CorreoUsuario <> "" Then
        m_Command = p_CorreoUsuario
    End If
    If m_Command <> "" Then
        Set m_ObjUsuarioConectado = Constructor.getUsuario(, , , m_Command, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set objNetwork = CreateObject("Wscript.Network")
        m_UsuarioLogeadoEnOrdenador = objNetwork.UserName
        If m_UsuarioLogeadoEnOrdenador = "Local1" Then m_UsuarioLogeadoEnOrdenador = "adm"
        If m_UsuarioLogeadoEnOrdenador = "adm1" Then m_UsuarioLogeadoEnOrdenador = "adm"
        Set m_ObjUsuarioConectado = Constructor.getUsuario(, m_UsuarioLogeadoEnOrdenador, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set objNetwork = Nothing
    End If
    If m_ObjUsuarioConectado Is Nothing Then
        p_Error = "No se ha podido determinar el usuario que está usando la herramienta"
        Err.Raise 1000
    End If
    ' issue-55: guardar contra usuarios dados de baja. Reguladorio: un usuario
    ' con FechaBaja < Date no debe poder iniciar sesion. Antes EVE no consultaba
    ' Usuario.Activo, asi que un ex-empleado con FechaBaja vencida podia entrar
    ' y ver/editar datos.
    If m_ObjUsuarioConectado.Activo = EnumSiNo.No Then
        p_Error = "EVE: el usuario '" & m_ObjUsuarioConectado.UsuarioRed & "' esta dado de baja (FechaBaja=" & _
            m_ObjUsuarioConectado.FechaBaja & "). Acceso denegado por seguridad."
        Err.Raise 1000
    End If
    If m_ObjUsuarioConectadoInicialmente Is Nothing Then
        Set m_ObjUsuarioConectadoInicialmente = m_ObjUsuarioConectado
    End If
    If m_ObjEntorno.ColUsuariosAdministradores.Exists(m_ObjUsuarioConectadoInicialmente.UsuarioRed) Then
        EsAdministradorConectadoInicialmente = EnumSiNo.Sí
    Else
        EsAdministradorConectadoInicialmente = EnumSiNo.No
    End If
    EsAdministrador = m_ObjEntorno.UsuarioConectadoEsAdministrador
    If EsAdministrador <> EnumSiNo.Sí Then
        
        EsCalidad = m_ObjEntorno.UsuarioConectadoEsDeCalidad
        If EsCalidad <> EnumSiNo.Sí Then
            EsTecnico = EnumSiNo.Sí
        Else
            EsTecnico = EnumSiNo.No
        End If
    End If
    
    
    t2 = Timer
    If EsAdministrador = EnumSiNo.Sí Then
        VBA.DoEvents
        'Debug.Print "EsAdministrador:Sí"
        VBA.DoEvents
        EsCalidad = EnumSiNo.No
        EsTecnico = EnumSiNo.No
    End If
    If EsCalidad = EnumSiNo.Sí Then
        VBA.DoEvents
        'Debug.Print "EsCalidad:Sí"
        VBA.DoEvents
        EsAdministrador = EnumSiNo.No
        EsTecnico = EnumSiNo.No
    Else
        VBA.DoEvents
        'Debug.Print "EsTecnico:Sí"
        VBA.DoEvents
    End If
    
    t2 = Timer
    
    VBA.DoEvents
    'Debug.Print "CargarUsuario: " & m_ObjUsuarioConectado.Nombre & vbTab & "T:" & t2 - t1
    VBA.DoEvents
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_EnOficina = EnumSiNo.No
    Else
        m_EnOficina = EnOficina(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
  
    
    For Each m_NombreCampo In m_ObjEntorno.ColItems.keys
        'Debug.Print m_nombreCampo
        'If CStr(m_nombreCampo) = "ColUsuarios" Then Stop
        Avance m_NombreCampo
        m_TipoCampo = m_ObjEntorno.ColItems(m_NombreCampo)
        If m_TipoCampo = "o" Then
            Set m_Objeto = m_ObjEntorno.getPropiedad(m_NombreCampo, p_Error)
            If p_Error <> "" Then
                If m_CadenaCamposConError = "" Then
                    m_CadenaCamposConError = m_NombreCampo
                Else
                    m_CadenaCamposConError = m_CadenaCamposConError & vbNewLine & m_NombreCampo
                End If
                intNumeroErrores = intNumeroErrores + 1
                p_Error = ""
            End If
            
        Else
            m_ValorCampo = m_ObjEntorno.getPropiedad(m_NombreCampo, p_Error)
            If p_Error <> "" Then
                If m_CadenaCamposConError = "" Then
                    m_CadenaCamposConError = m_NombreCampo
                Else
                    m_CadenaCamposConError = m_CadenaCamposConError & vbNewLine & m_NombreCampo
                End If
                intNumeroErrores = intNumeroErrores + 1
                p_Error = ""
            End If
            m_ValorCampoTruncado = Left(m_ValorCampo, 10) & " ..."
            
        End If
    Next
    t1 = Timer
    Avance "Cargando Perfiles"
    
    
    
    
    If m_ObjEntorno.VerSoloRiesgosNoRetirados = Empty Then
        m_ObjEntorno.VerSoloRiesgosNoRetirados = EnumSiNo.Sí
    End If
    If m_ObjEntorno.VerRiesgosDescripcion = Empty Then
        m_ObjEntorno.VerRiesgosDescripcion = EnumSiNo.Sí
    End If
    Set m_ObjUsuarioParaTareas = Nothing
    
   
    
    tf = Timer
    VBA.DoEvents
    Debug.Print "EVE en ......." & tf - ti
    If intNumeroErrores > 0 Then
        p_Error = "Se han producido los siguientes Errores: " & vbNewLine & m_CadenaCamposConError
        Err.Raise 1000
    End If
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo EVE ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
    Debug.Print p_Error
    
End Function


Public Function LeerIni(key As String, Default As Variant) As String
    Dim bufer As String * 256, Len_Value As Long
    
    
    Len_Value = GetPrivateProfileString(fso.GetBaseName(CurrentDb().Name), _
                                         key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         m_ObjEntorno.URLAchivoIni)
    LeerIni = Left$(bufer, CLng(Len_Value))
    
End Function



Public Function GetIPAddresses(Optional FilterLocalhost As Boolean = False) As String

    Dim Ret As Long
    Dim Buffer() As Byte
    Dim IPTableRow As IPINFO
    Dim Count As Long
    Dim BufferRequired As Long
    Dim StructSize As Long
    Dim NumIPAddresses As Long
    Dim IPAddress As String

  
        
    Call GetIpAddrTable(ByVal 0&, BufferRequired, 1)

    If BufferRequired > 0 Then
        
        ReDim Buffer(0 To BufferRequired - 1) As Byte
        
        If GetIpAddrTable(Buffer(0), BufferRequired, 1) = 0 Then
        
            'We've successfully obtained the IP Address details...
            'First 4 bytes is a long indicating the number of entries in the table
            StructSize = LenB(IPTableRow)
            CopyMemory NumIPAddresses, Buffer(0), 4
        
            While Count < NumIPAddresses
            
                'Buffer contains the IPINFO structures (after initial 4 byte long)
                CopyMemory IPTableRow, Buffer(4 + (Count * StructSize)), StructSize
                    
                IPAddress = IPAddressToString(IPTableRow.dwAddr)
                    
                If Not ((IPAddress = "127.0.0.1") _
                        And FilterLocalhost) Then
                            
                    'Replace this with whatever you want to do with the IP Address...
                    GetIPAddresses = GetIPAddresses & IPAddress & ";     "
                        
                End If
                
                Count = Count + 1
                
            Wend
            
        End If
            
    End If
 
    Exit Function



End Function

Private Function ParentPath(ByVal p_URL As String) As String
    If p_URL = "" Then
        ParentPath = ""
        Exit Function
    End If
    ParentPath = fso.GetParentFolderName(p_URL)
End Function

' ResetGlobals — limpia el cache de backend y de entidades
Public Sub ResetGlobals(ByRef p_Error As String)
    On Error GoTo errores
    If Not m_CachedDB Is Nothing Then
        On Error Resume Next
        m_CachedDB.Close
        On Error GoTo errores
    End If
    Set m_CachedDB = Nothing
    m_DBOpen = False
    m_ActiveBackendURL = ""
    m_PasswordBackend = ""
    m_BackendSandboxURL = ""
    m_IDAplicacion = ""
    IDAplicacion = ""
    m_EnPruebas = False
    m_RutaDirApp = ""
    m_RutaDirApp_PROD = ""
    m_RutaDirApp_LOCAL = ""
    ' Limpiar caches de entidades
    Set m_DicRiesgos = Nothing
    Set m_DicEdiciones = Nothing
    Set m_DicProyectos = Nothing
    Set m_DicPMs = Nothing
    Set m_DicPCs = Nothing
    Exit Sub
errores:
    p_Error = "ResetGlobals: " & Err.Number & " - " & Err.Description
End Sub

' ResetGetDbCache — cierra SOLO el cache de getdb (m_CachedDB) y resetea m_DBOpen.
' NO toca m_ActiveBackendURL, m_BackendSandboxURL, m_BackendActivo ni TempVars.
' Usado por tests para garantizar inter-test isolation per access-vba-tdd skill §1.7
' (sin esto, un handle cerrado del test anterior filtra al siguiente y rompe el
' cache check del branch non-testing de getdb).
Public Sub ResetGetDbCache()
    On Error Resume Next
    If Not m_CachedDB Is Nothing Then
        m_CachedDB.Close
    End If
    Set m_CachedDB = Nothing
    m_DBOpen = False
    On Error GoTo 0
End Sub

' Normaliza ruta: asegura que termine en backslash
Private Function NormalizarRuta(ByVal sRuta As String) As String
    If sRuta = "" Then
        NormalizarRuta = ""
    ElseIf Right(sRuta, 1) = "\" Then
        NormalizarRuta = sRuta
    Else
        NormalizarRuta = sRuta & "\"
    End If
End Function

' SanitizarRutaUsuarioWindowsLocal — adapta rutas locales C:\Users\<usuario> al perfil actual.
' El segundo parámetro permite tests deterministas sin depender de la cuenta Windows real.
Public Function SanitizarRutaUsuarioWindowsLocal(ByVal sRuta As String, Optional ByVal sPerfilUsuarioActual As String = "") As String
    Const USERS_PREFIX As String = "C:\Users\"

    Dim sPerfil As String
    Dim sPerfilNormalizado As String
    Dim lngPosSeparadorUsuario As Long

    If Len(sRuta) = 0 Then
        SanitizarRutaUsuarioWindowsLocal = ""
        Exit Function
    End If

    If StrComp(Left$(sRuta, Len(USERS_PREFIX)), USERS_PREFIX, vbTextCompare) <> 0 Then
        SanitizarRutaUsuarioWindowsLocal = sRuta
        Exit Function
    End If

    lngPosSeparadorUsuario = InStr(Len(USERS_PREFIX) + 1, sRuta, "\", vbTextCompare)
    If lngPosSeparadorUsuario = 0 Then
        SanitizarRutaUsuarioWindowsLocal = sRuta
        Exit Function
    End If

    sPerfil = sPerfilUsuarioActual
    If Len(sPerfil) = 0 Then
        sPerfil = Environ$("USERPROFILE")
    End If

    If Len(sPerfil) = 0 Then
        SanitizarRutaUsuarioWindowsLocal = sRuta
        Exit Function
    End If

    sPerfilNormalizado = sPerfil
    If Right$(sPerfilNormalizado, 1) = "\" Then
        sPerfilNormalizado = Left$(sPerfilNormalizado, Len(sPerfilNormalizado) - 1)
    End If

    SanitizarRutaUsuarioWindowsLocal = sPerfilNormalizado & Mid$(sRuta, lngPosSeparadorUsuario)
End Function

' Test_EVE — inicializador de testing.
' Fuerza BackendActivo=LOCAL antes de cargar la configuración, así getdb() siempre
' apunta al backend sandbox/local independientemente de lo que diga TbConfiguracionBackends.
' Luego llama a ResetGlobals + LeeConfiguracionLocal con el override aplicado.
Public Sub Test_EVE(ByRef p_Error As String)
    Dim sError As String
    p_Error = ""
    
    ' 1. Reset limpio para descartar cualquier estado cacheado de getdb()
    ResetGlobals sError
    If sError <> "" Then
        p_Error = "Test_EVE: Error en ResetGlobals: " & sError
        Exit Sub
    End If
    
    ' 2. Forzar modo local para que LeeConfiguracionLocal cargue desde BackendSandbox
    m_BackendActivo = "LOCAL"

    ' 3. Cargar configuración — m_BackendSandboxURL y paths locales se pueblan aquí
    LeeConfiguracionLocal sError
    If sError <> "" Then
        p_Error = "Test_EVE: Error en LeeConfiguracionLocal: " & sError
        Exit Sub
    End If

    ' 4. Forzar getdb() a que apunte SIEMPRE al backend sandbox/local en testing.
    '    Esto sobrescribe lo que haya leído LeeConfiguracionLocal de BackendActivo.
    m_BackendActivo = "LOCAL"
    m_ActiveBackendURL = m_BackendSandboxURL
    '    Invalida cache de getdb() para que la próxima llamadaabra con la URL correcta
    m_DBOpen = False
End Sub

' LeeConfiguracionLocal — lee de TbConfiguracionBackends
Public Sub LeeConfiguracionLocal(ByRef p_Error As String)
    Dim dbTemp As DAO.Database
    Dim rcdConfig As DAO.Recordset
    Dim sSQL As String
    
    On Error GoTo errores
    
    Set dbTemp = CurrentDb
    sSQL = "SELECT TOP 1 * FROM TbConfiguracionBackends"
    Set rcdConfig = dbTemp.OpenRecordset(sSQL, dbOpenSnapshot)
    
    Dim backendActivo As String
    If Not rcdConfig.EOF Then
        backendActivo = Nz(rcdConfig.Fields("BackendActivo").Value, "PROD")
        m_BackendActivo = backendActivo
        m_BackendSandboxURL = SanitizarRutaUsuarioWindowsLocal(Nz(rcdConfig.Fields("BackendSandbox").Value, ""))  ' siempre disponible para tests
        If backendActivo = "PROD" Then
            m_ActiveBackendURL = Nz(rcdConfig.Fields("BackendProduccion").Value, "")  ' archivo .accdb — NO normalizar
            m_RutaDirApp_PROD = NormalizarRuta(Nz(rcdConfig.Fields("RutaDirectorioAplicacion_PROD").Value, ""))
            m_RutaDirApp = m_RutaDirApp_PROD
        Else
            m_ActiveBackendURL = m_BackendSandboxURL  ' archivo .accdb — NO normalizar
            m_RutaDirApp_LOCAL = NormalizarRuta(SanitizarRutaUsuarioWindowsLocal(Nz(rcdConfig.Fields("RutaDirectorioAplicacion_LOCAL").Value, "")))
            m_RutaDirApp = m_RutaDirApp_LOCAL
        End If
        ' Dual-write TempVar (REQ-RES-004). Mirrors m_ActiveBackendURL into the
        ' contract source of truth that getdb() reads in non-testing mode.
        Application.TempVars("BackendPathConfigurado") = m_ActiveBackendURL
        m_PasswordBackend = Nz(rcdConfig.Fields("PasswordBackend").Value, "")
        m_IDAplicacion = Nz(rcdConfig.Fields("IDAplicacion").Value, 1)
        ' URLRutaAplicaciones: directorio PADRE del directorio de aplicación
        If m_RutaDirApp_PROD <> "" Then
            m_URLRutaAplicacionesRemotas = NormalizarRuta(fso.GetParentFolderName(m_RutaDirApp_PROD))
        End If
        If m_RutaDirApp_LOCAL <> "" Then
            m_URLRutaAplicacionesLocal = NormalizarRuta(fso.GetParentFolderName(m_RutaDirApp_LOCAL))
        End If
        ' EnPruebas: campo texto "Sí"/"No" — lectura defensiva por si no existe aún
        Dim fieldEnPruebas As DAO.Field
        Dim valorEnPruebas As String
        On Error Resume Next
        Set fieldEnPruebas = rcdConfig.Fields("EnPruebas")
        If Err.Number = 0 Then
            valorEnPruebas = Nz(fieldEnPruebas.Value, "No")
            m_EnPruebas = (valorEnPruebas = "Sí")
        Else
            m_EnPruebas = False
        End If
        On Error GoTo errores
        If m_EnPruebas Then
            Application.TempVars("EnPruebas") = "Sí"
        Else
            Application.TempVars("EnPruebas") = "No"
        End If
        ' --- TempVars config ---
        m_CadenaJerarquicaModelo = Nz(rcdConfig.Fields("CadenaJerarquicaModelo").Value, "nuevo")
        Application.TempVars("CadenaJerarquicaModelo") = m_CadenaJerarquicaModelo
        m_JPMesesAvisoEntreEdiciones = Nz(rcdConfig.Fields("JPMesesAvisoEntreEdiciones").Value, 3)
        Application.TempVars("JPMesesAvisoEntreEdiciones") = m_JPMesesAvisoEntreEdiciones
        m_JPDiasPreviosParaElAviso = Nz(rcdConfig.Fields("JPDiasPreviosParaElAviso").Value, 15)
        Application.TempVars("JPDiasPreviosParaElAviso") = m_JPDiasPreviosParaElAviso
        m_CalDiaInicialMesAviso = Nz(rcdConfig.Fields("CalDiaInicialMesAviso").Value, 2)
        Application.TempVars("CalDiaInicialMesAviso") = m_CalDiaInicialMesAviso
        m_Publicabilidad_Usar_Cache = Nz(rcdConfig.Fields("Publicabilidad_Usar_Cache").Value, "No")
        Application.TempVars("Publicabilidad_Usar_Cache") = m_Publicabilidad_Usar_Cache
        IDAplicacion = m_IDAplicacion
    Else
        m_ActiveBackendURL = ""
        ' Dual-write TempVar (REQ-RES-004, EOF branch) so a cleared
        ' m_ActiveBackendURL propagates to the contract source of truth.
        Application.TempVars("BackendPathConfigurado") = m_ActiveBackendURL
        m_PasswordBackend = ""
        m_BackendSandboxURL = ""
        m_IDAplicacion = ""
        m_EnPruebas = False
        m_RutaDirApp = ""
        m_RutaDirApp_PROD = ""
        m_RutaDirApp_LOCAL = ""
        m_CadenaJerarquicaModelo = "nuevo"
        m_JPMesesAvisoEntreEdiciones = 3
        m_JPDiasPreviosParaElAviso = 15
        m_CalDiaInicialMesAviso = 2
        m_Publicabilidad_Usar_Cache = "No"
        Application.TempVars("EnPruebas") = "No"
        Application.TempVars("CadenaJerarquicaModelo") = m_CadenaJerarquicaModelo
        Application.TempVars("JPMesesAvisoEntreEdiciones") = m_JPMesesAvisoEntreEdiciones
        Application.TempVars("JPDiasPreviosParaElAviso") = m_JPDiasPreviosParaElAviso
        Application.TempVars("CalDiaInicialMesAviso") = m_CalDiaInicialMesAviso
        Application.TempVars("Publicabilidad_Usar_Cache") = m_Publicabilidad_Usar_Cache
        IDAplicacion = m_IDAplicacion
    End If
    
    rcdConfig.Close
    Set rcdConfig = Nothing
    Set dbTemp = Nothing
    
    Exit Sub
errores:
    p_Error = "LeeConfiguracionLocal: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not rcdConfig Is Nothing Then rcdConfig.Close
    If Not dbTemp Is Nothing Then Set dbTemp = Nothing
End Sub

' getdb — devuelve el database del backend con cache Condor.
' Patrón alineado con 00_NO_CONFORMIDADES_staging:
'   - Testing mode: usa m_BackendSandboxURL directamente (nunca m_ActiveBackendURL)
'   - Non-testing mode: resuelve URL via Application.TempVars("BackendPathConfigurado")
'     con fallback a LeeConfiguracionLocal cuando el TempVar está vacío.
'   - p_SkipConfigLoad: salta LeeConfiguracionLocal (usa TempVar tal cual; error si está vacío)
Public Function getdb(Optional ByRef p_Error As String, _
                     Optional ByVal p_SkipConfigLoad As Boolean = False) As DAO.Database
    On Error GoTo errores
    p_Error = ""

    ' === TESTING MODE: sandbox directo, NUNCA m_ActiveBackendURL ===
    If m_TestingMode Then
        Dim sandboxPath As String
        sandboxPath = Trim$(m_BackendSandboxURL)

        If sandboxPath = "" Then
            If Not m_CachedDB Is Nothing Then
                On Error Resume Next
                m_CachedDB.Close
                Set m_CachedDB = Nothing
                m_DBOpen = False
                On Error GoTo errores
            End If
            p_Error = "getdb: m_TestingMode=True pero m_BackendSandboxURL esta vacio. No se abre backend activo/produccion."
            Set getdb = Nothing
            Exit Function
        End If

        If Not m_CachedDB Is Nothing Then
            Dim cachedPath As String
            Dim cachedErr As Long
            On Error Resume Next
            cachedPath = m_CachedDB.Name
            cachedErr = Err.Number
            On Error GoTo errores

            If cachedErr <> 0 Or StrComp(cachedPath, sandboxPath, vbTextCompare) <> 0 Then
                On Error Resume Next
                m_CachedDB.Close
                Set m_CachedDB = Nothing
                m_DBOpen = False
                On Error GoTo errores
            End If
        End If

        Dim fsoSandbox As Object
        Set fsoSandbox = CreateObject("Scripting.FileSystemObject")
        If Not fsoSandbox.FileExists(sandboxPath) Then
            Set fsoSandbox = Nothing
            If Not m_CachedDB Is Nothing Then
                On Error Resume Next
                m_CachedDB.Close
                Set m_CachedDB = Nothing
                m_DBOpen = False
                On Error GoTo errores
            End If
            p_Error = "getdb: m_TestingMode=True pero BackendSandbox no existe: " & sandboxPath
            Set getdb = Nothing
            Exit Function
        End If
        Set fsoSandbox = Nothing

        If Not m_CachedDB Is Nothing Then
            m_DBOpen = True
            Set getdb = m_CachedDB
            Exit Function
        End If

        On Error Resume Next
        Set m_CachedDB = DBEngine(0).OpenDatabase(sandboxPath, False, False, ";PWD=" & m_PasswordBackend)
        If Err.Number <> 0 Then
            p_Error = "getdb: m_TestingMode=True pero BackendSandbox no se pudo abrir: " & Err.Number & " - " & Err.Description
            Err.Clear
            Set m_CachedDB = Nothing
            m_DBOpen = False
            Set getdb = Nothing
            On Error GoTo errores
            Exit Function
        End If
        On Error GoTo errores

        m_DBOpen = True
        Set getdb = m_CachedDB
        Exit Function
    End If
    
    If m_DBOpen And Not m_CachedDB Is Nothing Then
        On Error Resume Next
        Dim dummy As String: dummy = m_CachedDB.Name  ' Error #3420 if stale
        On Error GoTo errores
        Set getdb = m_CachedDB
        Exit Function
    End If
    
    If Not m_CachedDB Is Nothing Then
        On Error Resume Next
        m_CachedDB.Close
        Set m_CachedDB = Nothing
        On Error GoTo errores
    End If
    
    ' Resolver URL desde TempVar (NO desde m_ActiveBackendURL)
    Dim m_URL As String
    m_URL = Nz(Application.TempVars("BackendPathConfigurado"), "")

    If m_URL = "" And Not p_SkipConfigLoad Then
        Call LeeConfiguracionLocal(p_Error)
        If p_Error <> "" Then Err.Raise 1000
        m_URL = Nz(Application.TempVars("BackendPathConfigurado"), "")
    End If

    If m_URL = "" Then
        p_Error = "BackendPathConfigurado no esta resuelto. Revise BackendActivo/BackendProduccion/BackendSandbox en TbConfiguracionBackends"
        Set getdb = Nothing
        Exit Function
    End If

    Dim ws As DAO.Workspace
    Set ws = DBEngine(0)
    Set m_CachedDB = ws.OpenDatabase(m_URL, False, False, ";PWD=" & m_PasswordBackend)
    m_DBOpen = True

    Set getdb = m_CachedDB
    Exit Function
errores:
    If Err.Number = 3420 Or Err.Number = 3024 Then
        ResetGlobals p_Error
        LeeConfiguracionLocal p_Error
        If p_Error = "" Then
            Set getdb = getdb(p_Error, p_SkipConfigLoad)
        Else
            Set getdb = Nothing
        End If
    Else
        If Err.Number <> 1000 Then p_Error = "getdb: " & Err.Number & " - " & Err.Description
        Set getdb = Nothing
    End If
End Function

Private Function IPAddressToString(EncodedAddress As Long) As String
    Dim IPBytes(3) As Byte
    Dim Count As Long
    
    'Converts a long IP Address to a string formatted 255.255.255.255
    'Note: Could use inet_ntoa instead
    
    CopyMemory IPBytes(0), EncodedAddress, 4 ' IP Address is stored in four bytes (255.255.255.255)
    
    'Convert the 4 byte values to a formatted string
    While Count < 4
        
        IPAddressToString = IPAddressToString & _
                                CStr(IPBytes(Count)) & _
                                IIf(Count < 3, ".", "")

        Count = Count + 1
            
    Wend
        
End Function



