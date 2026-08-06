Attribute VB_Name = "Variables Globales"
Option Compare Database
Option Explicit

Public Enum Aplicaciones
    AgedysEco = 2
    AGEDYSTEC = 3
    AGEDO = 4
    RIESGOS = 5
    BRASS = 6
    NC = 8
    Registro = 10
    SEGURIDAD = 16
End Enum
Public Enum EnumSino
    Sí = 1
    No = 2
End Enum
Public Enum EnumAplicaciones
    AgedysEco = 2
    AGEDYSTEC = 3
    AGEDO = 4
    GESTIONRIESGOS = 5
    BRASS = 6
    NoConformidades = 7
    Registro = 8
    Lanzadera = 12
    SEGURIDAD = 16
    HPS = 17
End Enum
Public Enum EnumPassEstados
    Correcto = 1
    Incorrecto = 2
    NecesitaCambio = 3
    caducada = 4
    Bloqueada = 5
End Enum
Public Enum EnumPermisos
    EsUsuarioAdministrador = 1
    EsUsuarioCalidad = 2
    EsUsuarioEconomia = 3
    EsUsuarioSecretaria = 4
    EsUsuarioTecnico = 5
    EsUsuarioSinAcceso = 6
    EsUsuarioCalidadAvisos = 7
End Enum
Public Enum EnumObjetos
    Aplicacion = 1
    Conexion = 2
    Correo = 3
    HistoricoPass = 4
    usuario = 5
    UsuarioAplicacionPermisos = 6
End Enum
Public Enum EnumVideoCategoria
    Ayuda = 1
    Encuesta = 2
End Enum

Public Enum EnumWMPPlayState
    wmppsStopped = 1
    wmppsPaused = 2
    wmppsPlaying = 3
    wmppsScanForward = 4
    wmppsScanReverse = 5
    wmppsBuffering = 6
    wmppsWaiting = 7
    wmppsMediaEnded = 8
    wmppsTransitioning = 9
    wmppsReady = 10
    wmppsReconnecting = 11
    wmppsLast = 12
End Enum

Public Enum EnumApertura
    Todas = 1
    HoyCerradas = 2
    HoyTodas = 3
    HoyAbiertas = 4
    TodasAbiertas = 5
End Enum

Public Const xlGeneral As Long = 1
Public Const xlBottom As Long = -4107
Public Const xlContext As Long = -5002

Public Const xlTypePDF As Long = 0
Public Const msoFileDialogFilePicker As Long = 3
Public Const msoFileDialogFolderPicker As Long = 4
Public Const msoFileDialogOpen As Long = 1
Public Const msoFileDialogSaveAs As Long = 2

Public m_UsuarioRedConectado As String
Public m_CorreoUsuarioEnLogin As String
Public m_SQL As String
Public Pregunta As Long
Public dato As Variant
Public m_ObjEntorno As Entorno

Public m_ObjUsuarioActivo As usuario
Public m_ObjAplicacionActiva As Aplicacion
Public m_ObjVideoActivo As Video
Public NodoActivo As MSComctlLib.Node
Public m_Linea As String


Public wks As DAO.Workspace
Private m_CachedDB As DAO.Database
Private m_DBOpen As Boolean
Private m_ActiveBackendURL As String
Private m_BackendProduccion As String
Private m_BackendSandbox As String
Private m_PasswordBackend As String
Private m_TestBackendURL As String
Public m_BackendActivo As String
Public m_EnPruebas As EnumSino
Public m_TestingMode As Boolean
Public m_RutaDirApp_PROD As String
Public m_RutaDirApp_LOCAL As String

Public IDAplicacion As String


Public intVecesEVE As Integer

Public EntradoComoOtro As EnumSino
Public HechoLogout As EnumSino
Public m_EnOficina As EnumSino
Public m_AperturaSinLogin As Boolean
'Public m_SSID As String
'Public m_SSIDPorArchivo As String
Public m_CorreoPorArchivo As String
'Public m_FechaPorArchivo As String
Public m_ObjUsuarioConectadoLogin As usuario

Public fso As New Scripting.FileSystemObject
Public Const SubRedOficina As String = "10.14.7"

Public EVEEjecutado As Boolean
Public lbl As Label
Public m_URLRutaAplicacionesLocal As String
Public m_URLRutaAplicacionesRemotas As String
Public m_URLRutaAplicacionLocal As String
Public m_URLRutaAplicacionRemota As String

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

Public Function getNombreUsuarioConectado(Optional ByRef p_Error As String) As String
    
    
    Dim m_UsuarioMaquina As usuario
    
    
    On Error GoTo errores
    If Not m_ObjUsuarioConectadoLogin Is Nothing Then
        getNombreUsuarioConectado = m_ObjUsuarioConectadoLogin.Nombre
        Exit Function
    End If
    If m_ObjEntorno Is Nothing Then
        getNombreUsuarioConectado = "Desconocido"
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

Public Function getdb( _
                        Optional p_EnumBBDDDatos As EnumAplicaciones, _
                        Optional ByRef p_Error As String _
                        ) As DAO.Database

    Dim intRetry As Integer
    Dim dummyName As String
    Dim m_TargetBackendURL As String
    On Error GoTo errores

retry_open:
    If m_TestingMode Then
        If m_TestBackendURL = "" Then
            If m_BackendSandbox = "" Then
                LeeConfiguracionLocal p_Error
                If p_Error <> "" Then Err.Raise 1000
            End If
            m_TestBackendURL = m_BackendSandbox
        End If

        m_TargetBackendURL = Trim$(m_TestBackendURL)
        If m_TargetBackendURL = "" Then
            p_Error = "getdb: m_TestingMode=True pero no existe un BackendSandbox configurado"
            Err.Raise 1000
        End If
    Else
        If m_ActiveBackendURL = "" Then
            LeeConfiguracionLocal p_Error
            If p_Error <> "" Then Err.Raise 1000
        End If

        m_TargetBackendURL = Trim$(m_ActiveBackendURL)
    End If

    If m_DBOpen And Not m_CachedDB Is Nothing Then
        On Error Resume Next
        dummyName = m_CachedDB.Name
        If Err.Number = 0 Then
            If StrComp(dummyName, m_TargetBackendURL, vbTextCompare) = 0 Then
                On Error GoTo errores
                Set getdb = m_CachedDB
                Exit Function
            End If

            m_CachedDB.Close
        End If
        Set m_CachedDB = Nothing
        m_DBOpen = False
        On Error GoTo errores
    End If

    If Not fso.FileExists(m_TargetBackendURL) Then
        p_Error = "No es alcanzable la base de datos " & vbNewLine & m_TargetBackendURL
        Err.Raise 1000
    End If

    Set wks = DBEngine.Workspaces(0)
    Set m_CachedDB = wks.OpenDatabase(m_TargetBackendURL, False, False, "MS Access;PWD=" & m_PasswordBackend)
    m_DBOpen = True
    Set getdb = m_CachedDB
    Exit Function
errores:
    If Err.Number = 3420 And intRetry = 0 Then
        intRetry = intRetry + 1
        p_Error = ""
        ResetGlobals p_Error
        p_Error = ""
        GoTo retry_open
    End If
    If Err.Number <> 1000 Then
        p_Error = "El método getdb ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function ResetGlobals(Optional ByRef p_Error As String) As String
    Dim m_KeepTestingMode As Boolean
    Dim m_KeepTestBackendURL As String

    On Error GoTo errores
    p_Error = ""

    m_KeepTestingMode = m_TestingMode
    m_KeepTestBackendURL = m_TestBackendURL

    If m_DBOpen And Not m_CachedDB Is Nothing Then
        On Error Resume Next
        m_CachedDB.Close
        On Error GoTo errores
    End If

    Set m_CachedDB = Nothing
    m_DBOpen = False
    m_ActiveBackendURL = ""
    m_BackendProduccion = ""
    m_BackendSandbox = ""
    m_PasswordBackend = ""
    m_BackendActivo = ""
    m_EnPruebas = EnumSino.No
    m_RutaDirApp_PROD = ""
    m_RutaDirApp_LOCAL = ""
    m_TestingMode = m_KeepTestingMode
    If m_TestingMode Then
        m_TestBackendURL = m_KeepTestBackendURL
    Else
        m_TestBackendURL = ""
    End If

    ResetGlobals = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ResetGlobals ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function LeeConfiguracionLocal( _
                                    Optional ByRef p_Error As String, _
                                    Optional ByVal p_BackendActivoForzado As String = "" _
                                    ) As String
    Dim dbLocal As DAO.Database
    Dim rst As DAO.Recordset
    Dim m_EnPruebasValor As String

    On Error GoTo errores
    p_Error = ""

    Set dbLocal = CurrentDb
    Set rst = dbLocal.OpenRecordset("SELECT TOP 1 * FROM TbConfiguracionBackends", dbOpenSnapshot)
    If rst.EOF Then
        p_Error = "TbConfiguracionBackends no tiene registros"
        Err.Raise 1000
    End If

    m_BackendActivo = UCase$(Trim$(Nz(rst.Fields("BackendActivo").value, "")))
    m_BackendProduccion = Trim$(Nz(rst.Fields("BackendProduccion").value, ""))
    m_BackendSandbox = Trim$(Nz(rst.Fields("BackendSandbox").value, ""))
    m_PasswordBackend = Nz(rst.Fields("PasswordBackend").value, "")
    If m_PasswordBackend = "" Then
        m_PasswordBackend = Environ$("ACCESS_VBA_PASSWORD")
    End If
    IDAplicacion = CStr(Nz(rst.Fields("IDAplicacion").value, ""))
    m_RutaDirApp_PROD = NormalizarRuta(CStr(Nz(rst.Fields("RutaDirectorioAplicacion_PROD").value, "")))
    m_RutaDirApp_LOCAL = NormalizarRuta(CStr(Nz(rst.Fields("RutaDirectorioAplicacion_LOCAL").value, "")))
    m_EnPruebasValor = UCase$(Trim$(Nz(rst.Fields("EnPruebas").value, "NO")))

    m_EnPruebas = IIf(m_EnPruebasValor = "SI" Or m_EnPruebasValor = "SÍ", EnumSino.Sí, EnumSino.No)

    p_BackendActivoForzado = UCase$(Trim$(p_BackendActivoForzado))
    If p_BackendActivoForzado <> "" Then
        m_BackendActivo = p_BackendActivoForzado
    End If

    m_BackendActivo = UCase$(Trim$(m_BackendActivo))
    If m_BackendActivo = "" Then
        p_Error = "BackendActivo es obligatorio en TbConfiguracionBackends"
        Err.Raise 1000
    End If

    Select Case m_BackendActivo
        Case "LOCAL"
            m_ActiveBackendURL = m_BackendSandbox
        Case "PROD"
            m_ActiveBackendURL = m_BackendProduccion
        Case Else
            p_Error = "BackendActivo inválido en TbConfiguracionBackends: " & m_BackendActivo
            Err.Raise 1000
    End Select

    If m_ActiveBackendURL = "" Then
        p_Error = "No existe una URL de backend activa válida"
        Err.Raise 1000
    End If
    If IDAplicacion = "" Then
        p_Error = "IDAplicacion es obligatorio en TbConfiguracionBackends"
        Err.Raise 1000
    End If

    TempVarSet "BackendActivo", m_BackendActivo
    TempVarSet "BackendProduccion", m_BackendProduccion
    TempVarSet "BackendSandbox", m_BackendSandbox
    TempVarSet "IDAplicacion", IDAplicacion
    TempVarSet "EnPruebas", IIf(m_EnPruebas = EnumSino.Sí, "Sí", "No")
    TempVarSet "RutaDirectorioAplicacion", ResolverRutaSegunConfiguracion(m_RutaDirApp_LOCAL, m_RutaDirApp_PROD, p_Error)

    m_URLRutaAplicacionesRemotas = m_RutaDirApp_PROD
    m_URLRutaAplicacionesLocal = m_RutaDirApp_LOCAL
    m_URLRutaAplicacionRemota = m_RutaDirApp_PROD
    m_URLRutaAplicacionLocal = m_RutaDirApp_LOCAL

    LeeConfiguracionLocal = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método LeeConfiguracionLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function SetTestingMode( _
                            ByVal p_Enabled As Boolean, _
                            Optional ByVal p_TestBackendURL As String = "", _
                            Optional ByRef p_Error As String = "" _
                            ) As String
    On Error GoTo errores
    p_Error = ""

    m_TestingMode = p_Enabled
    If p_Enabled Then
        m_TestBackendURL = Trim$(p_TestBackendURL)
    Else
        m_TestBackendURL = ""
    End If

    SetTestingMode = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SetTestingMode ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function GetTestingBackendURL() As String
    GetTestingBackendURL = m_TestBackendURL
End Function

Public Function BackendActivoEsLocal(Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo errores
    p_Error = ""

    If m_BackendActivo = "" Then
        LeeConfiguracionLocal p_Error
        If p_Error <> "" Then Err.Raise 1000
    End If

    BackendActivoEsLocal = (UCase$(Trim$(m_BackendActivo)) = "LOCAL")
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método BackendActivoEsLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function EntornoEnPruebasConfigurado(Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo errores
    p_Error = ""

    If m_BackendActivo = "" Then
        LeeConfiguracionLocal p_Error
        If p_Error <> "" Then Err.Raise 1000
    End If

    EntornoEnPruebasConfigurado = (m_EnPruebas = EnumSino.Sí)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EntornoEnPruebasConfigurado ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function ResolverRutaSegunConfiguracion( _
                                        ByVal p_RutaLocal As String, _
                                        ByVal p_RutaProduccion As String, _
                                        Optional ByRef p_Error As String = "" _
                                        ) As String
    On Error GoTo errores
    p_Error = ""

    If BackendActivoEsLocal(p_Error) Then
        ResolverRutaSegunConfiguracion = p_RutaLocal
    Else
        If p_Error <> "" Then Err.Raise 1000
        ResolverRutaSegunConfiguracion = p_RutaProduccion
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ResolverRutaSegunConfiguracion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function NombreAplicacionSegunConfiguracion(Optional ByRef p_Error As String = "") As String
    On Error GoTo errores
    p_Error = ""

    If EntornoEnPruebasConfigurado(p_Error) Then
        NombreAplicacionSegunConfiguracion = "LANZADERA PRUEBAS"
    Else
        If p_Error <> "" Then Err.Raise 1000
        NombreAplicacionSegunConfiguracion = "LANZADERA"
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método NombreAplicacionSegunConfiguracion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function GetBackendPassword(Optional ByRef p_Error As String) As String
    On Error GoTo errores
    p_Error = ""

    If m_PasswordBackend = "" Then
        LeeConfiguracionLocal p_Error
        If p_Error <> "" Then Err.Raise 1000
    End If

    If m_PasswordBackend = "" Then
        p_Error = "No existe contraseña de backend configurada"
        Err.Raise 1000
    End If

    GetBackendPassword = m_PasswordBackend
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GetBackendPassword ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function EVE( _
                    Optional ByRef p_Error As String, _
                    Optional ByVal p_BackendActivoForzado As String = "" _
                    ) As String
    
    Dim m_NombreCampo As Variant
    Dim m_Valor As String
    Dim m_Objeto As Object
    
    Dim ti As Single
    Dim tf As Single
    Dim m_TipoCampo As String
    Dim m_ValorCampo As String
    Dim m_ValorCampoTruncado As String
    Dim intNumeroErrores As Integer
    Dim m_CadenaCamposConError As String
    On Error GoTo errores
    
    ti = Timer
    
    Set m_ObjUsuarioConectadoLogin = Nothing

    ResetGlobals p_Error
    If p_Error <> "" Then Err.Raise 1000
    LeeConfiguracionLocal p_Error, p_BackendActivoForzado
    If p_Error <> "" Then Err.Raise 1000

    m_EnOficina = EnOficina(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjEntorno Is Nothing Then
        Set m_ObjEntorno = New Entorno
    End If
    
    For Each m_NombreCampo In m_ObjEntorno.ColItems.Keys
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
    
    
    tf = Timer
    If intNumeroErrores > 0 Then
        p_Error = "Se han producido los siguientes Errores: " & vbNewLine & m_CadenaCamposConError
        Err.Raise 1000
    End If
    EVEEjecutado = True
    
    
    
    Debug.Print "Tareas en ......." & tf - ti
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EVE ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    EVEEjecutado = False
    Debug.Print p_Error
End Function

Public Function EVEMin( _
                        Optional ByRef p_Error As String _
                        ) As String
    
   
    
    
    On Error GoTo errores
    
    Application.TempVars.RemoveAll
    ResetGlobals p_Error
    If p_Error <> "" Then Err.Raise 1000
    LeeConfiguracionLocal p_Error
    If p_Error <> "" Then Err.Raise 1000

    Set m_ObjEntorno = New Entorno
    m_AperturaSinLogin = AperturaSinLogin(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioConectadoLogin = Nothing
    m_CorreoUsuarioEnLogin = ""
    m_CorreoPorArchivo = getCorreoPorSSID(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_AperturaSinLogin = True Then
       
        Set m_ObjUsuarioConectadoLogin = Constructor.getUsuario(p_Correo:=m_CorreoPorArchivo, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ObjUsuarioConectadoLogin Is Nothing Then
            m_CorreoUsuarioEnLogin = m_ObjUsuarioConectadoLogin.CorreoUsuario
            If Not EVEEjecutado Then
                Avance "Estableciendo las variables de entorno"
                EVE p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
            Set m_ObjUsuarioConectadoLogin = Constructor.getUsuario(p_Correo:=m_CorreoPorArchivo, p_Error:=p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EVEMin ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
   
    Debug.Print p_Error
End Function

Public Function LoginCorrecto(Optional ByRef p_Error As String) As String
    
    Dim m_NombreFormulario As String
    On Error GoTo errores
    
    
    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
        m_NombreFormulario = "FormMenuPrincipalAdmin"
    ElseIf m_ObjUsuarioConectadoLogin.EsUsuarioCalidadCalculado = EnumSino.Sí Then
        m_NombreFormulario = "FormMenuPrincipalCalidad"
    Else
        m_NombreFormulario = "FormMenuPrincipalUsuario"
    End If
            
    If FormularioAbierto(m_NombreFormulario) Then
        DoCmd.Close acForm, m_NombreFormulario, acSaveNo
    End If
    DoCmd.Close acForm, "FormLogin", acSaveNo
    DoCmd.OpenForm m_NombreFormulario
    Exit Function
errores:
    If Err.Number <> "" Then
        p_Error = "El método LoginCorrecto ha devuelto el error " & vbNewLine & Err.Description
    End If
End Function

Public Function DameURLEjecutableAccess(Optional ByRef p_Error As String) As String
    
    
    Dim m_ArchivosDePrograma As String
    Dim m_URLEjecutable As String
    Dim intOrdinal As Integer
    On Error GoTo errores
    
    m_ArchivosDePrograma = Environ("CommonProgramFiles")
    If InStr(1, m_ArchivosDePrograma, "\Common Files") <> 0 Then
        m_ArchivosDePrograma = Replace(m_ArchivosDePrograma, "\Common Files", "")
    Else
        m_ArchivosDePrograma = Replace(m_ArchivosDePrograma, "\Archivos comunes", "")
    End If
    If Right(m_ArchivosDePrograma, 1) <> "\" Then
        m_ArchivosDePrograma = m_ArchivosDePrograma & "\"
    End If
    For intOrdinal = 11 To 20
        m_URLEjecutable = m_ArchivosDePrograma & "Microsoft Office\OFFICE" & CStr(intOrdinal) & "\MSACCESS.EXE"
        If fso.FileExists(m_URLEjecutable) Then
            DameURLEjecutableAccess = m_URLEjecutable
            Exit Function
        End If
    Next
    m_URLEjecutable = m_ArchivosDePrograma & "Microsoft Office\root\Office16" & "\MSACCESS.EXE"
    If fso.FileExists(m_URLEjecutable) Then
        DameURLEjecutableAccess = m_URLEjecutable
        Exit Function
    End If
    
    DameURLEjecutableAccess = m_URLEjecutable
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameURLEjecutableAccess ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    
End Function

Public Function ColIconosDeAccesoNoAlcanzables(Optional ByRef p_Error As String) As Scripting.Dictionary
    
    
    Dim m_objAplicacion As Aplicacion
    Dim m_ID As Variant
    Dim m_URLAcceso As String
    Dim m_URLNoAcceso As String
    Dim m_mensaje As String
    On Error GoTo errores
    If m_ObjEntorno.ColAplicaciones Is Nothing Then
        p_Error = m_ObjEntorno.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    For Each m_ID In m_ObjEntorno.ColAplicaciones
        
        m_mensaje = ""
        Set m_objAplicacion = m_ObjEntorno.ColAplicaciones(m_ID)
        If m_objAplicacion.ConIconoEnLanzaderaCalculado = EnumSino.No Then
            GoTo siguiente
        End If
        
        m_URLAcceso = m_objAplicacion.URLIconoAcceso
        p_Error = m_objAplicacion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not fso.FileExists(m_URLAcceso) Then
            m_mensaje = "Acceso " & m_objAplicacion.NombreAplicacion & ":" & m_URLAcceso
        End If
        m_URLNoAcceso = m_objAplicacion.URLIconoNOAcceso
        p_Error = m_objAplicacion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not fso.FileExists(m_URLNoAcceso) Then
            If m_mensaje = "" Then
                m_mensaje = "NO Acceso " & m_objAplicacion.NombreAplicacion & ":" & m_URLNoAcceso
            Else
                m_mensaje = m_mensaje & "|" & "NO Acceso " & m_objAplicacion.NombreAplicacion & ":" & m_URLNoAcceso
            End If
        End If
        If m_mensaje <> "" Then
            If ColIconosDeAccesoNoAlcanzables Is Nothing Then
                Set ColIconosDeAccesoNoAlcanzables = New Scripting.Dictionary
            End If
            ColIconosDeAccesoNoAlcanzables.CompareMode = TextCompare
            If Not ColIconosDeAccesoNoAlcanzables.Exists(CStr(m_ID)) Then
                ColIconosDeAccesoNoAlcanzables.Add CStr(m_ID), m_mensaje
            End If
        End If
siguiente:
        Set m_objAplicacion = Nothing
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ColIconosDeAccesoNoAlcanzables ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    
End Function

Public Function NormalizarRuta(ByVal p_Ruta As String) As String
    Dim m_Ruta As String
    m_Ruta = Trim$(p_Ruta)
    If m_Ruta = "" Then Exit Function
    m_Ruta = Replace(m_Ruta, "/", "\\")
    If Right$(m_Ruta, 1) <> "\" Then m_Ruta = m_Ruta & "\"
    NormalizarRuta = m_Ruta
End Function

Private Function TempVarSet(ByVal p_Nombre As String, ByVal p_Valor As Variant) As String
    On Error Resume Next
    Application.TempVars.Remove p_Nombre
    On Error GoTo 0
    Application.TempVars.Add p_Nombre, p_Valor
End Function




