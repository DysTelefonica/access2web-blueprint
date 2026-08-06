Attribute VB_Name = "Variables Globales"
Option Compare Database
Option Explicit

#If Win64 = 1 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" ( _
            ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
            ByVal lpApplicationName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As String, _
            ByVal nSize As Long, _
            ByVal lpFileName As String) As Long
    
    Public Declare PtrSafe Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hwnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As Long)
    Public Declare PtrSafe Function GetIpAddrTable Lib "Iphlpapi" ( _
            pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
    Public Declare PtrSafe Function IsIconic Lib "user32.dll" ( _
            ByVal hwnd As Long) As Long
    Public Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" ( _
        ByVal hProcess As Long, lpExitCode As Long) As Long
    Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    Public Declare PtrSafe Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameA" ( _
        ByVal lpszShortPath As String, _
        ByVal lpszLongPath As String, _
        ByVal cchBuffer As Long) As Long
    Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Public Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
   
#Else
    Public Declare Sub Sleep Lib "kernel32" ( _
            ByVal dwMilliseconds As Long)
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
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As Long)
    Public Declare Function GetIpAddrTable Lib "IPHlpApi" ( _
            pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
    Public Declare Function IsIconic Lib "user32.dll" ( _
            ByVal hWnd As Long) As Long
    Public Declare Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Public Declare Function GetExitCodeProcess Lib "kernel32" ( _
        ByVal hProcess As Long, lpExitCode As Long) As Long
    Public Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    Public Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameA" ( _
        ByVal lpszShortPath As String, _
        ByVal lpszLongPath As String, _
        ByVal cchBuffer As Long) As Long
    Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function EmptyClipboard Lib "user32" () As Long
    Public Declare Function CloseClipboard Lib "user32" () As Long
    Public Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As Long) As Long
    Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    
#End If

Public Const GMEM_MOVEABLE = &H2
Public Const CF_TEXT = 1
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STATUS_PENDING = &H103&
Public Const msoFileDialogFilePicker As Long = 3
Public Const msoFileDialogFolderPicker As Long = 4
Public Const msoFileDialogOpen As Long = 1
Public Const msoFileDialogSaveAs As Long = 2
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
    no = 2
End Enum

Public Enum EnumEstadoGeneral
    EnTramitacion = 1
    Tramitado = 2
End Enum
Public Enum EnumEstado
   
    PendienteEnvioExcel = 2
    PendienteRecordatorioExcel1 = 3
    PendienteRecordatorioExcel2 = 4
    PendienteExcelRelleno = 5
    PendienteRegistroMarga = 6
    PendienteRecordatorioEnvioDPS1 = 7
    PendienteRecordatorioEnvioDPS2 = 8
    PendienteEnvioDPS = 9
    PendienteRegistroONS = 10
    PendienteRegistroEnHPS = 11
    Tramitado = 12
    Desestimado = 13
    Cancelado = 14
    PendienteEnvioTraspasoONS = 15
    EnTramite = 16
End Enum
Public Enum EnumPlantillas
    EnvioExcel = 1
    RecordatorioExcel1 = 2
    RecordatorioExcel2 = 3
    RecordatorioMarga1 = 4
    RecordatorioMarga2 = 5
    CancelacionPreMarga = 6
    CancelacionMarga = 7
    
End Enum
Public Enum EnumTipoEnvioCorreo
    EnvioExcel = 1
    RecordatorioExcel1 = 2
    RecordatorioExcel2 = 3
    RecordatorioMarga1 = 4
    RecordatorioMarga2 = 5
    CancelacionPreMarga = 6
    CancelacionMarga = 7
    AltaEnMarga = 8
    ExcelAdjuntado = 9
    EnvioDPS = 10
    EnvioONS = 11
    Desestimado = 12
    
    cancelacion = 14
End Enum
Public Enum EnumTipo
    Alta = 1
    Traspaso = 2
    Renovacion = 3
    ElevacionGrado = 4
End Enum
Public Enum EnumEntidades
    Empresas = 1
    Responsables = 2
    PlantillasHTML = 3
    MotivosHPS = 4
    JustificacionHPS = 5
End Enum
    Public fso As New Scripting.FileSystemObject
    Public Const SubRedOficina As String = "10.14.7"
    Public m_SQL As String
    Public pregunta As Long
    Public DescParaLog As String
    Public lbl As Label
    Public wks As DAO.Workspace
    Private DB As DAO.Database
    Private db1 As DAO.Database
    Public IDAplicacion As String
    Public m_ObjUsuarioConectado As Usuario
    Public m_ObjEntorno As Entorno
    Public m_EnOficina As EnumSiNo
    Public EsAdministrador As EnumSiNo
    Public EsTecnico As EnumSiNo
    Public m_FaltaAlgunParametroEnConfiguracion As Boolean
    Public m_ObjConfiguracion As Configuracion
    Public m_TituloFormulario As String
    Public m_ObjSolicitudActiva As solicitud
    Public m_ObjSolicitudAlInicio As solicitud
    Public m_ObjMotivoActivo As MotivoHPS
    Public m_ObjTramitadorParaTareas As Usuario
    
    Public m_ObjExpedienteActivo As Expediente
    Public m_ObjCorreoActivo As Correo
    Public m_ObjResposableActivo As Responsable
    Public m_ObjUsuarioHPSActivo As UsuarioHPS
    Public m_URLHTMLActivo As String
    Public m_URLRutaAplicacionesLocal As String
    Public m_URLRutaAplicacionesRemotas As String
    Public m_URLRutaAplicacionLocal As String
    Public m_URLRutaAplicacionRemota As String
    Public Const CONST_IDSOLICITUD_PRUEBA As String = "-1"
    
    
    
    
Private Function VaciarVariables()
   
    Set m_ObjSolicitudActiva = Nothing
    Set m_ObjExpedienteActivo = Nothing
    Set m_ObjConfiguracion = Nothing
    m_URLHTMLActivo = ""
    Set m_ObjUsuarioConectado = Nothing
    m_TituloFormulario = ""
    m_EnOficina = Empty
    EsAdministrador = Empty
    EsTecnico = Empty
    Set m_ObjTramitadorParaTareas = Nothing
End Function


Public Function EVE(Optional ByRef p_Error As String) As String
    
    Dim m_NombreCampo As Variant
    Dim m_Valor As String
    Dim m_Objeto As Object
    Dim ti As Single
    Dim tf As Single
    Dim m_TipoCampo As String
    Dim m_ValorCampo As String
    Dim intNumeroErrores As Integer
    Dim m_CadenaCamposConError As String
   
    
    Dim objNetwork As Object
    Dim m_Command As String
    Dim m_UsuarioLogeadoEnOrdenador As String
    Dim t1 As Single
    Dim t2 As Single
    Dim m_UsuarioDeRed As String
    On Error GoTo errores
    
    
    t1 = Timer
    VaciarVariables
    
    Application.TempVars.RemoveAll
    Application.TempVars("EnDesarrollo") = "No"
    'Application.TempVars("EnDesarrollo") = "Sí"
    Application.TempVars("DatosEnLocal") = "No"
    'Application.TempVars("DatosEnLocal") = "Sí"
    'Application.TempVars("EnPruebas") = "Sí"
    Application.TempVars("EnPruebas") = "No"
    Application.TempVars("ConCorreoCopiaGestor") = "Sí"
    Application.TempVars("ActivadoCorreoAutomatico") = "Sí"
    'Application.TempVars("ActivadoCorreoAutomatico") = "No"
    Application.TempVars("RegistroEnHPS") = "Sí"
    'Application.TempVars("RegistroEnHPS") = "No"
    
    Application.TempVars("ExpedienteUnificado") = "No"
    'Application.TempVars("ExpedienteUnificado") = "Sí"
    'm_EnOficina = EnOficina(p_Error)
    Avance "Estableciendo variables de entorno"
   
    IDAplicacion = "22"
   
    Set m_ObjUsuarioConectado = Nothing
    m_URLRutaAplicacionesRemotas = "\\datoste\aplicaciones_dys\Aplicaciones PpD\"
    m_URLRutaAplicacionRemota = m_URLRutaAplicacionesRemotas & "SOLICITUDES HPS\"
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URLRutaAplicacionesLocal = getRutaAplicacionesLocal(p_Error)
        If m_URLRutaAplicacionesLocal <> "" Then
            m_URLRutaAplicacionLocal = m_URLRutaAplicacionesLocal & "SOLICITUDES HPS\"
        End If
    Else
       m_URLRutaAplicacionesLocal = ""
       m_URLRutaAplicacionLocal = ""
    End If
    
     m_Command = Nz(VBA.Command, "")
    'm_Command = "emma.delgadillogomez@telefonica.com"
    'm_Command = "martina.torralbarodriguez@telefonica.com"
    'm_Command = "angel.martin-doradocaballero@telefonica.com"
    'm_Command = "esperanza.delalamoarriba@telefonica.com"
    If m_Command <> "" Then
        Set m_ObjUsuarioConectado = constructor.getUsuario(, , , m_Command, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set objNetwork = CreateObject("Wscript.Network")
        m_UsuarioLogeadoEnOrdenador = objNetwork.UserName
        If m_UsuarioLogeadoEnOrdenador = "Local1" Then m_UsuarioLogeadoEnOrdenador = "adm"
        If m_UsuarioLogeadoEnOrdenador = "adm1" Then m_UsuarioLogeadoEnOrdenador = "adm"
        Set m_ObjUsuarioConectado = constructor.getUsuario(, m_UsuarioLogeadoEnOrdenador, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set objNetwork = Nothing
    End If
    
    
    If m_ObjUsuarioConectado Is Nothing Then
        p_Error = "No se ha podido determinar el usuario que está usando la herramienta"
        Err.Raise 1000
    End If
    Set m_ObjEntorno = New Entorno
     With m_ObjUsuarioConectado
        EsAdministrador = .EsAdministradorCalculado
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If EsAdministrador <> EnumSiNo.Sí Then
            EsTecnico = .EsUsuarioTecnicoCalculado
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If EsTecnico <> EnumSiNo.Sí Then
                p_Error = "Usuario no autorizado"
                Err.Raise 1000
            End If
           
        Else
            
        End If
    End With
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_EnOficina = EnumSiNo.no
    Else
        m_EnOficina = EnOficina(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    'm_EnOficina = EnumSiNo.Sí
            
    For Each m_NombreCampo In m_ObjEntorno.ColItems.Keys
        'Debug.Print m_NombreCampo
        'If CStr(m_NombreCampo) = "DiasParaAutocancelacion" Then Stop
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
            
            
        End If
    Next
    
    
    If intNumeroErrores > 0 Then
        p_Error = "Se han producido los siguientes Errores: " & vbNewLine & m_CadenaCamposConError
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        If Application.TempVars("EnPruebas") = "Sí" Then
            m_TituloFormulario = "SOLICITUDES HPS Versión " & m_ObjEntorno.VersionAplicacion & " " & _
                    m_ObjUsuarioConectado.Nombre & " (Sólo lectura) EN PRUEBAS"
        Else
             m_TituloFormulario = "SOLICITUDES HPS Versión " & m_ObjEntorno.VersionAplicacion & " " & _
                    m_ObjUsuarioConectado.Nombre & " (Sólo lectura)"
        End If
    Else
        If Application.TempVars("EnPruebas") = "Sí" Then
            m_TituloFormulario = "SOLICITUDES HPS Versión " & m_ObjEntorno.VersionAplicacion & " " & _
                m_ObjUsuarioConectado.Nombre & " EN PRUEBAS"
        Else
             m_TituloFormulario = "SOLICITUDES HPS Versión " & m_ObjEntorno.VersionAplicacion & " " & _
                m_ObjUsuarioConectado.Nombre
        End If
    End If
    Set m_ObjConfiguracion = m_ObjEntorno.Configuracion
    
    If m_ObjConfiguracion Is Nothing Then
        p_Error = "No se ha podido obtener la configuración"
        Err.Raise 1000
    End If
    m_FaltaAlgunParametroEnConfiguracion = DatosConfInsuficientes(m_ObjConfiguracion, EnumSiNo.no, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    
    
    
    t2 = Timer
    Debug.Print "Variables establecidas correctamente en: " & t2 - t1
    AvanceCerrar
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EVE ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function getdbLanzadera( _
                                Optional ByRef p_Error As String _
                                ) As DAO.Database
    
    Dim m_URL As String
    On Error GoTo errores
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\Lanzadera_Datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "0Lanzadera\Lanzadera_Datos.accdb"
    Else
        p_Error = "No se conoce el origen de los datos"
        Err.Raise 1000
    End If
   
    
    
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
    
     If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\Solicitudes_HPS_datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionRemota & "Solicitudes_HPS_datos.accdb"
    Else
        p_Error = "No se conoce el origen de los datos"
        Err.Raise 1000
    End If
   
    
    
    Set wks = DBEngine.Workspaces(0)
    Set DB = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    
    Set getdb = DB
    
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
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\Correos_datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "00Recursos\Correos_datos.accdb"
    Else
        p_Error = "No se sabe si se está usando en local o en remoto"
        Err.Raise 1000
    End If
    
       
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

Public Function getdbHPS( _
                            Optional ByRef p_Error As String _
                            ) As DAO.Database
    
    Dim m_URL As String
    
    On Error GoTo errores
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\HPST.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "HPS\HPST.accdb"
    Else
        p_Error = "No se sabe si se está usando en local o en remoto"
        Err.Raise 1000
    End If
    
    
    
    Set wks = DBEngine.Workspaces(0)
    Set DB = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    
    Set getdbHPS = DB
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbHPS ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getNombreUsuarioConectado(Optional ByRef p_Error As String) As String
    
    Dim m_Usuario As Usuario
    Dim m_UsuarioMaquina As Usuario
    Dim m_Nombre As String
    
    On Error GoTo errores
    If Not m_ObjUsuarioConectado Is Nothing Then
        getNombreUsuarioConectado = m_ObjUsuarioConectado.Nombre
        Exit Function
    End If
    
    Set m_UsuarioMaquina = constructor.getUsuarioConectadoPorMaquina(p_Error)
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



Public Function getdbExpedientes( _
                                Optional ByRef p_Error As String _
                                ) As DAO.Database
    
    Dim m_URL As String
    On Error GoTo errores
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\Expedientes_datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "EXPEDIENTES\Expedientes_datos.accdb"
    Else
        p_Error = "No se sabe si se está usando en local o en remoto"
        Err.Raise 1000
    End If
    
    
    
    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbExpedientes = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbExpedientes ha devuelto el error: " & vbNewLine & Err.Description
    End If
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
