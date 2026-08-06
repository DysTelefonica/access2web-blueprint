Attribute VB_Name = "VariablesEntorno"
Option Compare Database
Option Explicit
#If Win64 = 1 Then
    Public Declare PtrSafe Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
            ByVal lpApplicationName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As String, _
            ByVal nSize As Long, _
            ByVal lpFileName As String) As Long
    Public Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" ( _
        ByVal hProcess As Long, lpExitCode As Long) As Long
    Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As LongPtr
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
    Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
    Public Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
    Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As LongPtr)
    Public Declare PtrSafe Function GetIpAddrTable Lib "Iphlpapi" ( _
            pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
#Else
    Public Declare Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
            ByVal lpApplicationName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As String, _
            ByVal nSize As Long, _
            ByVal lpFileName As String) As Long
    
    
   
    Public Declare Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Public Declare Function GetExitCodeProcess Lib "kernel32" ( _
        ByVal hProcess As Long, lpExitCode As Long) As Long
    Public Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Public Declare Function CloseClipboard Lib "user32" () As Long
    Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function EmptyClipboard Lib "user32" () As Long
    Public Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare  Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As Long)
    Public Declare  Function GetIpAddrTable Lib "Iphlpapi" ( _
            pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
#End If

Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STATUS_PENDING = &H103&

'The structures returned by the API call GetIpAddrTable...
Public Type IPINFO
    dwAddr As Long          ' IP address
    dwIndex As Long         ' interface index
    dwMask As Long          ' subnet mask
    dwBCastAddr As Long     ' broadcast address
    dwReasmSize  As Long    ' assembly size
    Reserved1 As Integer
    Reserved2 As Integer
End Type

Public Enum EnumSiNo
    Sí = 1
    No = 2
End Enum
Public Enum EnumFiltroResultado
    Todos = 1
    DatosUsuario = 2
    HPSNAC = 3
    HPSOTAN = 4
    HPSUE = 5
    HPSESA = 6
End Enum
Public Enum EnumEstadoVisibleHPS
    Activo = 1 '--->APuntoCaducarRenovacionNo+APuntoCaducarRenovacionSiSinSolicitud+APuntoCaducarRenovacionSiConSolicitud
    ActivoAPuntoCaducar = 2 '---->APuntoCaducarRenovacionSiSinSolicitud+APuntoCaducarRenovacionSiConSolicitud
    Caducada = 3  '----->CaducadaRenovacionNo+CaducadaRenovacionSiSinSolicitud+CaducadaRenovacionSiConSolicitud
    Baja = 4
    
    Solicitada = 5 '---->ActivoSolicitada+ActivoSinSolicitar+APuntoCaducarRenovacionSiConSolicitud+CaducadaRenovacionSiConSolicitud+SolicitudNueva
    Irregular = 6
    SinOrdenDeRenovacion = 7 '---->
    NoPosee = 8
    PendienteRenovacion = 9 '---->APuntoCaducarRenovacionSiSinSolicitud
End Enum
Public Enum EnumEstadoHPS
    APuntoCaducarRenovacionNo = 1
    APuntoCaducarRenovacionSiSinSolicitud = 2
    APuntoCaducarRenovacionSiConSolicitud = 3
    CaducadaRenovacionNo = 4
    CaducadaRenovacionSiSinSolicitud = 5
    CaducadaRenovacionSiConSolicitud = 6
    ActivoSolicitada = 7
    ActivoSinSolicitar = 8
    ActivoRenovacionNo = 9
    SolicitudNueva = 10
    PendienteRenovacion = 11
    Baja = 12
    
    Irregular = 13
    NoPosee = 14
     
End Enum
Public Type Tipo_HPSEstado
    Estado As String
    MotivoIrregular As String
End Type
Public Enum EnumTipoHPS
    Nacional = 1
    OTAN = 2
    UE = 3
    ESA = 4
End Enum
Public Enum EnumConsultas
    Todo = 1
    PorJuridicaQueContrata = 2
    PorProyectoAsignado = 3
    PorNombre = 4
    PorTipoHPS = 5
    PorEmpresaDelUsuario = 6
    PorJuridicaQueTramitaHPS = 7
    PorEstadoHPS = 8
    PorMotivoHPS = 9
End Enum

Public Enum EnumModoApertura
    master = 1
    slave = 2
End Enum
Public Enum EnumTipoObjeto
    Consulta = 1
    Expediente = 2
    ExpedienteGrado = 3
    HPS = 4
    JuridicasContratacion = 5
    Observacion = 6
    Anexo = 7
    UsuarioHPS = 8
    tabla = 9
    DatosLocal = 10
    UsuarioSICA = 11
    AnexoSICA = 12
    UsuarioHistorico = 13
    ObservacionHistorica = 14
    AnexoHistorico = 15
    AnexoUsuarioHistorico = 16
    Empresa = 17
End Enum
Public Enum EnumTipoSQL
    Select_ = 1
    SelectParaListas = 2
    Delete = 3
    INSERT = 4
    Update = 5
    
End Enum

Public Enum EnumTipoIndicador
    PendientesCursoHPS = 1
    HPSAPuntoDeCaducar = 2
    HPSACaducadas = 3
    HPSASolicitandose = 4
    PtesPrimeraConvocatoriaCurso = 5
    PtesSegundaConvocatoriaCurso = 6
    PtesEnvioCorreoJefeSeguridadCurso = 7
End Enum


Public Const xlGeneral As Long = 1
Public Const xlBottom As Long = -4107
Public Const xlContext As Long = -5002

Public Const msoFileDialogFilePicker As Long = 3
Public Const msoFileDialogFolderPicker As Long = 4
Public Const msoFileDialogOpen As Long = 1
Public Const msoFileDialogSaveAs As Long = 2

Public Const m_ColorAzul As Long = 8210719
Public Const m_ColorBlanco As Long = 16777215


Public fso As New Scripting.FileSystemObject
Public Const SubRedOficina As String = "10.14.7"
Public pregunta As Long, dato, m_SQL As String, flag As String
Public m_ObjEntorno As entorno
Public m_ObjConn As ADODB.Connection


Public m_ObjUsuarioActivo As UsuarioHPS
Public m_ObjUsuarioSICAActivo As UsuarioSICA
Public m_ObjUsuarioHistoricoActivo As UsuarioHistorico


Public m_ObjObservacionActiva As Observacion
Public m_ObjObservacionHistoricaActiva As ObservacionHistorica

Public m_ObjAnexoUsuarioHPSActivo As AnexoUsuarioHPS
Public m_ObjAnexoUsuarioSICAActivo As AnexoUsuarioSICA
Public m_ObjAnexoUsuarioHistoricoActivo As AnexoUsuarioHistorico


Public m_ObjHPSNACActiva As HPS
Public m_ObjHPSOTANActiva As HPS
Public m_ObjHPSESAActiva As HPS
Public m_ObjHPSUEActiva As HPS

Public m_ObjSuministradorActivo As Suministrador

Public m_ObjUsuarioAlInicio As UsuarioHPS
Public m_ObjUsuarioSICAAlInicio As UsuarioSICA
Public m_ObjUsuarioHistoricoAlInicio As UsuarioHistorico
Public m_ObjHPSNACAlInicio As HPS
Public m_ObjHPSOTANAlInicio As HPS
Public m_ObjHPSESAAlInicio As HPS
Public m_ObjHPSUEAlInicio As HPS


Public m_ObjDatosLocalActivos As DatosLocal
Public m_ObjColErrores As Collection
Public m_ObjColParaFiltros As Scripting.Dictionary

Public m_ObjIndicadores As Indicador
Public m_EstadoConexion As String

Public m_CommandPorUsuario As String
Public blnUsuarioEncontrado As Boolean
Public wks As DAO.Workspace
Private db As DAO.Database
Private db1 As DAO.Database
Public IDAplicacion As String
Public m_ObjUsuarioConectado As Usuario
Public m_ObjUsuarioConectadoInicialmente As Usuario
Public m_EnOficina As EnumSiNo
Public m_AccesoADatosTE As EnumSiNo

Private m_Command As String 'se obtienen cuando se abre la base de datos desde la oficina (correo)
Public blnYaRegistradaEntrada As Boolean

Public lbl As Label
Public m_URLRutaAplicacionesLocal As String
Public m_URLRutaAplicacionesRemotas As String
Public m_URLRutaAplicacionLocal As String
Public m_URLRutaAplicacionRemota As String




Public Function LeerIni(key As String, Default As Variant) As String
    Dim bufer As String * 256
    Dim Len_Value As Long
    
    
    Len_Value = GetPrivateProfileString("HPS", _
                                         key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         m_ObjEntorno.URLIni)
          
    LeerIni = Left$(bufer, CLng(Len_Value))
    
End Function
'Función para escrbir datos en archivos INI.

Public Function EVE( _
                    Optional ByRef p_Error As String _
                    ) As String
    
    Dim m_Propiedad As Variant
    Dim m_ValorPropiedad As String
    Dim m_ObjValorPropiedad As Object
    Dim t1 As Double
    Dim t2 As Double
    Dim objNetwork As Object
    Dim m_UsuarioConectado As String
    Dim m_CorreoUsuarioConectado As String
    Dim m_UsuarioLogeadoEnOrdenador As String
    Dim m_ParteObjeto As String
    Dim NumeroErrores As Integer
    
   
    
    Dim m_Errores As String
    
    On Error GoTo errores
    p_Error = ""
    Avance "Obteniendo variables globales"
    
    
    t1 = Timer
    
   
   
    Application.TempVars("DatosEnLocal") = "No"
    'Application.TempVars("DatosEnLocal") = "Sí"
    Application.TempVars("MesesParaRenovacion") = 9
    m_EnOficina = EnOficina(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    'm_EnOficina = EnumSiNo.No
    Set m_ObjUsuarioConectado = Nothing
    If Nz(VBA.Command, "") <> "" Then
        m_AccesoADatosTE = EnumSiNo.Sí
    Else
        If m_EnOficina = EnumSiNo.Sí Then
            m_AccesoADatosTE = EnumSiNo.Sí
        Else
            If fso.FolderExists("\\datoste\aplicaciones_dys\Aplicaciones PpD") Then
                m_AccesoADatosTE = EnumSiNo.Sí
            Else
                m_AccesoADatosTE = EnumSiNo.No
            End If
        End If
        
    End If
    If m_AccesoADatosTE = EnumSiNo.No Then
        p_Error = "Se ha de estar conectado al menos al CARU"
        Err.Raise 1000
    End If
    ' [config-as-source-of-truth] Las constantes m_URLRutaAplicaciones*
    ' ya NO se siembran con paths hardcoded. Se derivan de
    ' TbConfiguracionHPS vía m_ObjEntorno.GetConfigOrFail en cada uso
    ' (EVE carga m_ObjEntorno más abajo, así que aquí sólo dejamos los
    ' campos vacíos para mantener compatibilidad con código legacy que
    ' todavía los lee en zonas no migradas).
    m_URLRutaAplicacionesRemotas = ""
    m_URLRutaAplicacionRemota = ""
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URLRutaAplicacionesLocal = getRutaAplicacionesLocal(p_Error)
        If m_URLRutaAplicacionesLocal <> "" Then
            m_URLRutaAplicacionLocal = m_URLRutaAplicacionesLocal & "HPS\"
        End If
    Else
       m_URLRutaAplicacionesLocal = ""
       m_URLRutaAplicacionLocal = ""
    End If
    m_Command = VBA.Command
    'm_Command = "anamaria.rubiocanales@telefonica.com"
    'm_Command = "esperanza.delalamoarriba@telefonica.com"
    'm_Command = "martina.torralbarodriguez@telefonica.com"
    'm_Command = "andres.romandelperal@telefonica.com"
    m_CorreoUsuarioConectado = m_Command
    Set m_ObjEntorno = New entorno
    ' PR B: pre-poblamos la cache de TbConfiguracionHPS en m_ObjEntorno
    ' ANTES del bucle de inicializacion de propiedades de abajo. Si lo
    ' hicieramos al final, los property Gets (URLCarpetaAnexos y
    ' URLCarpetaAnexosHistoricos) ya habrian quedado cacheados con las
    ' rutas hardcoded y la cache quedaria inutil. On Error Resume Next
    ' para no romper el arranque si la tabla aun no existe -- los
    ' property Gets haran fallback al comportamiento legacy.
    On Error Resume Next
    m_ObjEntorno.CargarConfiguracion
    On Error GoTo errores
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
    If m_ObjUsuarioConectadoInicialmente Is Nothing Then
        Set m_ObjUsuarioConectadoInicialmente = m_ObjUsuarioConectado
    End If

    
    IDAplicacion = "17"
    
    
    If m_ObjUsuarioConectado.EsAdministradorCalculado = EnumSiNo.No And _
        m_ObjUsuarioConectado.EsTecnicoCalculado = EnumSiNo.No Then
        p_Error = "El usuario no tiene permiso para abrir la aplicación"
        Err.Raise 1000
    End If
    
    For Each m_Propiedad In m_ObjEntorno.ColPropiedadesEntorno
'        VBA.DoEvents
        'Debug.Print m_Propiedad
'       If m_Propiedad = "ColExpedientes" Then Stop
        Avance "Obteniendo variables globales..." & m_Propiedad
        m_ParteObjeto = m_ObjEntorno.ColPropiedadesEntorno(m_Propiedad)
        If m_ParteObjeto = "o" Then
            Set m_ObjValorPropiedad = m_ObjEntorno.getValorPropiedad(CStr(m_Propiedad))
            p_Error = m_ObjEntorno.Error
            If p_Error <> "" Then
                NumeroErrores = NumeroErrores + 1
                If m_Errores = "" Then
                    m_Errores = p_Error
                Else
                    m_Errores = m_Errores & vbNewLine & p_Error
                End If
                p_Error = ""
            End If
           
           
           Set m_ObjValorPropiedad = Nothing
        Else
            
            m_ValorPropiedad = m_ObjEntorno.getValorPropiedad(CStr(m_Propiedad), p_Error)
            
            If p_Error <> "" Then
                NumeroErrores = NumeroErrores + 1
                If m_Errores = "" Then
                    m_Errores = p_Error
                Else
                    m_Errores = m_Errores & vbNewLine & p_Error
                End If
                p_Error = ""
            End If
            If m_ParteObjeto = "s" Then
                If m_ValorPropiedad = "1" Then
                    m_ValorPropiedad = "Sí"
                ElseIf m_ValorPropiedad = "2" Then
                    m_ValorPropiedad = "No"
                Else
                    m_ValorPropiedad = "#ERR"
                End If
            End If
            
            m_ValorPropiedad = ""
        End If
SiguientePropiedad:
        m_ObjEntorno.Error = ""
        
    Next
    If NumeroErrores > 0 Then
        p_Error = m_Errores
        Err.Raise 1000
    End If
    
    
    Set m_ObjUsuarioActivo = Nothing
    Set m_ObjUsuarioSICAActivo = Nothing
    Set m_ObjObservacionActiva = Nothing
    Set m_ObjAnexoUsuarioHPSActivo = Nothing
    Set m_ObjAnexoUsuarioSICAActivo = Nothing
    
    Set m_ObjHPSNACActiva = Nothing
    Set m_ObjHPSOTANActiva = Nothing
    Set m_ObjHPSESAActiva = Nothing
    Set m_ObjHPSUEActiva = Nothing
    t2 = Timer
    Debug.Print "Entorno establecido en : " & t2 - t1
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = m_Propiedad & vbTab & "El metodo EVE ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function ActualizarDatos( _
                                Optional ByRef p_Error As String _
                                ) As String
                
    Dim m_Propiedad As Variant
    Dim m_ValorPropiedad As String
    Dim m_ObjValorPropiedad As Object
    Dim t1 As Double
    Dim t2 As Double
    Dim objNetwork As Object
    Dim m_UsuarioConectado As String
    Dim m_CorreoUsuarioConectado As String
    Dim m_ParteObjeto As String
    Dim NumeroErrores As Integer
    
   
    
    Dim m_Errores As String
    
    On Error GoTo errores
    p_Error = ""
    'If p_DeActualizar = Empty Then p_DeActualizar = EnumSiNo.No
    
    
    t1 = Timer
    EVE p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
   
    ' Al usar la optimizada, NO hace falta regenerar (Sí), basta con sincronizar (No)
    InicializarTablasLocalesYEntidad_Optimizado p_ForzarRegeneracionTotal:=EnumSiNo.No, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set m_ObjUsuarioActivo = Nothing
    Set m_ObjUsuarioSICAActivo = Nothing
    Set m_ObjObservacionActiva = Nothing
    Set m_ObjAnexoUsuarioHPSActivo = Nothing
    Set m_ObjAnexoUsuarioSICAActivo = Nothing
   
    Set m_ObjHPSNACActiva = Nothing
    Set m_ObjHPSOTANActiva = Nothing
    Set m_ObjHPSESAActiva = Nothing
    Set m_ObjHPSUEActiva = Nothing
    t2 = Timer
    Debug.Print "ActualizarDatos establecido en : " & t2 - t1
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = m_Propiedad & vbTab & "El metodo ActualizarDatos ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function getdbLanzadera( _
                                Optional ByRef p_Error As String _
                                ) As DAO.Database
    
    Dim m_URL As String
    On Error GoTo errores
    
    ' [config-as-source-of-truth] La ruta del .accdb vive en
    ' TbConfiguracionHPS bajo la clave LANZADERA_BACKEND_PATH. GetConfigOrFail
    ' Err.Raise 1000 si la clave no está sembrada; el caller recibe un
    ' error accionable que apunta a ConfigurarTablaConfiguracion_PerfilLocal.
    m_URL = m_ObjEntorno.GetConfigOrFail("LANZADERA_BACKEND_PATH")

    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbLanzadera = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbLanzadera ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function getdb( _
                        Optional ByRef p_Error As String _
                        ) As DAO.Database
    
    Dim m_URL As String
    
    On Error GoTo errores

    ' [config-as-source-of-truth] HPST_BACKEND_PATH desde TbConfiguracionHPS.
    m_URL = m_ObjEntorno.GetConfigOrFail("HPST_BACKEND_PATH")
    
    Set wks = DBEngine.Workspaces(0)
    Set db = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    
    Set getdb = db
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdb ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function getdbCorreo( _
                                Optional ByRef p_Error As String _
                                ) As DAO.Database
    
    Dim m_URL As String
        
    On Error GoTo errores
    
    ' [config-as-source-of-truth] CORREOS_BACKEND_PATH desde TbConfiguracionHPS.
    m_URL = m_ObjEntorno.GetConfigOrFail("CORREOS_BACKEND_PATH")
    
    If Not fso.FileExists(m_URL) Then
        p_Error = "No se alcanza la URL de los datos: " & vbNewLine & m_URL
        Err.Raise 1000
    End If
    
    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbCorreo = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbCorreo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function getdbSolicitudHPS( _
                                    Optional ByRef p_Error As String _
                                    ) As DAO.Database
    
    Dim m_URL As String
    On Error GoTo errores
    
    ' [config-as-source-of-truth] SOLICITUDES_HPS_BACKEND_PATH desde
    ' TbConfiguracionHPS.
    m_URL = m_ObjEntorno.GetConfigOrFail("SOLICITUDES_HPS_BACKEND_PATH")
    
    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbSolicitudHPS = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbSolicitudHPS ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getdbExpedientes( _
                                Optional ByRef p_Error As String _
                                ) As DAO.Database
    
    Dim m_URL As String
    On Error GoTo errores
    
    ' [config-as-source-of-truth] EXPEDIENTES_BACKEND_PATH desde
    ' TbConfiguracionHPS.
    m_URL = m_ObjEntorno.GetConfigOrFail("EXPEDIENTES_BACKEND_PATH")
    
    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbExpedientes = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbExpedientes ha devuelto el error: " & vbNewLine & Err.Description
    End If
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

