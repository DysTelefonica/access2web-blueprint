Attribute VB_Name = "Funciones Genreales"
Option Compare Database
Option Explicit

#If Win64 = 1 Then
    Public Declare PtrSafe Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, ByVal lpParameters As String, _
            ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
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
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As LongPtr)
    Public Declare PtrSafe Function GetIpAddrTable Lib "Iphlpapi" ( _
            pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

#Else
    Public Declare Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, ByVal lpParameters As String, _
            ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
    Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
            ByVal lpBuffer As String, nSize As Long) As Long
    
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
Public Sub Ajustar(frmFormulario As Form)
    Dim i As Long
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    With frmFormulario
        ' ajusto el ancho del formulario teniendo en cuenta si tiene o no selector de registros
        If Not .RecordSelectors Then
            .InsideWidth = frmFormulario.Width
        Else
            .InsideWidth = frmFormulario.Width + 250
        End If
        ' si se abre en vista formulario simple
        If .DefaultView = 0 Then
           ' ajusto el alto incluyendo las distintas secciones, encabezado, pie, grupos...
           ' como no sé el número de secciones del formulario, me salgo al producirse un error
           .InsideHeight = 0
           For i = 0 To 100
              .InsideHeight = .InsideHeight + .Section(i).Height
           Next
        End If
    End With
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
End Sub

Function FicheroAbierto( _
                        m_URL As String, _
                        Optional ByRef p_Error As String _
                        ) As Boolean
    
    Dim intfilenum As Integer
    On Error GoTo errores
    intfilenum = FreeFile()
    Open m_URL For Binary Access Read Write Lock Read Write As #intfilenum
    Close #intfilenum
    If Err.Number <> 0 Then
        FicheroAbierto = True
    Else
        FicheroAbierto = False
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FicheroAbierto ha producido el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function EjecutarAplicacion( _
                                    p_URL As String, _
                                    lngHwn As Long, _
                                    Optional p_Error As String _
                                    ) As String
    
    Dim m_URLEjecutable As String
    
    On Error GoTo errores
    
    
    If Not fso.FileExists(p_URL) Then
        p_Error = "No se ha podido alcanzar el archivo"
        Err.Raise 1000
    End If
    m_URLEjecutable = m_ObjEntorno.URLEjecutableAccess
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URLEjecutable) Then
        p_Error = "No se puede obtener m_URLEjecutable"
        Err.Raise 1000
    End If
    
    
    ShellExecute lngHwn, "Open", m_URLEjecutable, """ & p_URL & """, "", 1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EjecutarAplicacion ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function AlterarPropriedadBBDD( _
                                        db As DAO.Database, _
                                        p_PropName As String, _
                                        p_PropType As Variant, _
                                        p_PropValue As Variant, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    '******************************************************************
    'Activa/Desactiva por software el acceso privilegiado con la tecla SHIFT
    'AlterarPropriedade "AllowBypassKey", dbBoolean, True --->>>>> Activa
    'AlterarPropriedade "AllowBypassKey", dbBoolean, False --->>>>> Desactiva
    'Solo se puede modificar desde el formulario de control --- Triple seguridad --- Pass de admin, usuario windows de admin, usuario BD Admin
    '*************************************************************************
    Dim m_prp As Property
    Const conPropNoEncontradaError As Long = 3270
    On Error Resume Next
    
    db.Properties(p_PropName) = p_PropValue
    If Err.Number = conPropNoEncontradaError Then
        Err.Clear
        On Error GoTo errores
        Set m_prp = db.CreateProperty(p_PropName, p_PropType, p_PropValue)
        db.Properties.Append m_prp
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AlterarPropriedadBBDD ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function LeerIni( _
                        key As String, _
                        Optional Default As Variant, _
                        Optional ByRef p_Error As String _
                        ) As String
   
    
    Dim Len_Value As Long
    Dim bufer As String * 256
    
    
    On Error GoTo errores
    
    
    
    If Not fso.FileExists(m_ObjEntorno.URLArchivoIniLanzadera) Then
        p_Error = m_ObjEntorno.Error
        If p_Error <> "" Then
            p_Error = "No se puede alcanzar el archivo de configuración" & vbNewLine & m_ObjEntorno.URLArchivoIniLanzadera
            Err.Raise 1000
        End If
    End If
    Len_Value = GetPrivateProfileString(fso.GetBaseName(CurrentProject.Name), _
                                         key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         m_ObjEntorno.URLArchivoIniLanzadera)
          
    LeerIni = Left$(bufer, CLng(Len_Value))
    
    Exit Function
errores:
    If p_Error <> "" Then
        p_Error = "El método LeerIni ha devuelto el error: " & Err.Description
    End If
    
End Function
Public Function getVersion( _
                            p_URLIni As String, _
                            Optional ByRef p_Error As String _
                            ) As String
   
    
    Dim Len_Value As Long
    Dim bufer As String * 256
    Dim key As String
    
    On Error GoTo errores
    
    
    
    If Not fso.FileExists(p_URLIni) Then
        Exit Function
    End If
    key = "VersionAplicacion"
    Len_Value = GetPrivateProfileString(fso.GetBaseName(p_URLIni), _
                                         key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         p_URLIni)
          
    getVersion = Left$(bufer, CLng(Len_Value))
    
    Exit Function
errores:
    If p_Error <> "" Then
        p_Error = "El método getVersion ha devuelto el error: " & Err.Description
    End If
    
End Function
Public Function LeerIniLocal( _
                                key As String, _
                                Optional Default As Variant, _
                                Optional ByRef p_Error As String _
                                ) As String
   
    
    Dim Len_Value As Long
    Dim bufer As String * 256
    Dim m_nombreBase As String
    Dim m_URL As String
    On Error GoTo errores
    
    
    m_URL = m_ObjEntorno.URLDirLocal & fso.GetBaseName(Application.CurrentProject.Name) & ".ini"
    If Not fso.FileExists(m_URL) Then
        
        Exit Function
    End If
    
    Len_Value = GetPrivateProfileString(fso.GetBaseName(Application.CurrentProject.Name), _
                                         key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         m_URL)
          
    LeerIniLocal = Left$(bufer, CLng(Len_Value))
    
    Exit Function
errores:
    If p_Error <> "" Then
        p_Error = "El método LeerIniLocal ha devuelto el error: " & Err.Description
    End If
    
End Function

Public Function EjecutarShell( _
                                strComando As String, _
                                Optional p_ManteniendoProceso As EnumSino = EnumSino.Sí, _
                                Optional ByRef p_Error As String _
                                ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 11/04/2020
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       Llama al shell de WSH permitiendo la posibilidad de hacerlo de manera sincrona o asincrona
    '       Argumentos: strURLRutaAplicacion => Ruta de la aplicación a ejecutar
    '                   strParametros     => (opcional) parámetros a pasarle a la aplic.
    '       uso: ShellWSH "C:\windows\notepad","C:\pp.txt", True
    '   -Llamada desde
    
    '   -Devuelve:
    '       EjecutarShell = Descriptivo
    
    '-------------------------------------------------------------------
    
    Dim ManejadorProceso As Long
    Dim IDProceso As Long
    Dim lpExitCode As Long
    On Error GoTo errores
    
    If strComando = "" Then
        p_Error = "No se ha indicado el comando"
        Err.Raise 1000
    End If
    IDProceso = Shell(strComando, vbHide)
    
    If p_ManteniendoProceso = EnumSino.Sí Then
        ManejadorProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, IDProceso)
        ' Mientras lp_ExitCode = STATUS_PENDING, se ejecuta el do
        Do
            Call GetExitCodeProcess(ManejadorProceso, lpExitCode)
            DoEvents
        Loop While lpExitCode = STATUS_PENDING
        Call CloseHandle(ManejadorProceso)
    End If
    
    
    EjecutarShell = "OK"
    Exit Function
errores:
    
    If Err.Number <> 1000 Then
        p_Error = "El método EjecutarShell ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function EjecutarShelllanzar( _
                                strComando As String, _
                                Optional p_ManteniendoProceso As EnumSino = EnumSino.Sí, _
                                Optional ByRef p_Error As String _
                                ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 11/04/2020
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       Llama al shell de WSH permitiendo la posibilidad de hacerlo de manera sincrona o asincrona
    '       Argumentos: strURLRutaAplicacion => Ruta de la aplicación a ejecutar
    '                   strParametros     => (opcional) parámetros a pasarle a la aplic.
    '       uso: ShellWSH "C:\windows\notepad","C:\pp.txt", True
    '   -Llamada desde
    
    '   -Devuelve:
    '       EjecutarShelllanzar = Descriptivo
    
    '-------------------------------------------------------------------
    
    Dim ManejadorProceso As Long
    Dim IDProceso As Long
    Dim lpExitCode As Long
    On Error GoTo errores
    
    If strComando = "" Then
        p_Error = "No se ha indicado el comando"
        Err.Raise 1000
    End If
    IDProceso = Shell(strComando, vbMaximizedFocus)
    
    If p_ManteniendoProceso = EnumSino.Sí Then
        ManejadorProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, IDProceso)
        ' Mientras lp_ExitCode = STATUS_PENDING, se ejecuta el do
        Do
            Call GetExitCodeProcess(ManejadorProceso, lpExitCode)
            DoEvents
        Loop While lpExitCode = STATUS_PENDING
        Call CloseHandle(ManejadorProceso)
    End If
    
    
    EjecutarShelllanzar = "OK"
    Exit Function
errores:
    
    If Err.Number <> 1000 Then
        p_Error = "El método EjecutarShelllanzar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function AbrirEnLocal( _
                                p_URLRemoto As String, _
                                lngHwnd As Long, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim m_URLLocal As String
    On Error GoTo errores
    
    
    If Not fso.FileExists(p_URLRemoto) Then
        p_Error = "No es accesible la ruta del archivo que se pretende abrir" & vbNewLine & p_URLRemoto
        Err.Raise 1000
    End If
    m_URLLocal = m_ObjEntorno.URLDirLocal & fso.GetFileName(p_URLRemoto)
    If fso.FileExists(m_URLLocal) Then
        If FicheroAbierto(m_URLLocal) Then
            p_Error = "Tiene el archivo abierto"
            Err.Raise 1000
        End If
    End If
    fso.CopyFile p_URLRemoto, m_URLLocal, True
    Ejecutar lngHwnd, "open", m_URLLocal, "", "", 1
    
    AbrirEnLocal = m_URLLocal
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function DameID( _
                        p_NOmbreTabla As String, _
                        p_NombreCampoID As String, _
                        Optional ByRef p_db As DAO.Database, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim lngIDMax As Long
    On Error GoTo errores
    
    If p_NOmbreTabla = "" Or p_NombreCampoID = "" Then
        p_Error = "Se ha de indicar el nombre de la tabla y de su campo ID"
        Err.Raise 1000
    End If
    If p_db Is Nothing Then
        Set p_db = getdb()
    End If
    m_SQL = "SELECT Max(" & p_NOmbreTabla & "." & p_NombreCampoID & ") AS MaxID " & _
            "FROM " & p_NOmbreTabla & ";"
    Set rcdDatos = p_db.OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MaxID"), "")) Then
                lngIDMax = .Fields("MaxID")
            End If

        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameID = CStr(lngIDMax + 1)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameID ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function FormularioAbierto( _
                                    strNombreFormulario As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Boolean
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día X
    '   -Modificaciones:
    '       -23/03/2011.- Formateo la función a como está la mayoría
    '   -Funcionamiento:
    '       -Va a mirar en la base de datos si está o no abierto el formulario
    '   -Llamada desde
    
    '   -Devuelve:
    '       FormularioAbierto = true or false
    
    '-------------------------------------------------------------------
    Dim strEstado As String
    On Error GoTo errores
    strEstado = SysCmd(acSysCmdGetObjectState, acForm, strNombreFormulario)
    If strEstado = "0" Then
        FormularioAbierto = False
    Else
        FormularioAbierto = True
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FormularioAbierto ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Public Function AbrirArchivoAyuda( _
                                    ByRef frm As Form, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_NombreAdjunto As String
    Dim m_URL As String
    
        
    On Error GoTo errores
    
    m_SQL = "SELECT TbHerramientaDocAyuda.NombreArchivoAyuda " & _
            "FROM TbHerramientaDocAyuda " & _
            "WHERE NombreFormulario='" & frm.Name & "' ;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            m_NombreAdjunto = Nz(.Fields("NombreArchivoAyuda"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If m_NombreAdjunto = "" Then
        Exit Function
    End If
    'm_URL = m_ObjEntorno.urldi & m_NombreAdjunto
    
    AbrirEnLocal m_URL, frm.hWnd, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    AbrirArchivoAyuda = m_URL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbrirArchivoAyuda ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Sub EjecuarAviso(Tipo As Long, strTitulo As String, strMensaje As String)
    Dim strURLIcono As String
End Sub


Public Function getSSIDporMaquina(Optional ByRef p_Error As String) As String
    
    Dim strComando As String
    Dim fichero As TextStream
    
    Dim strLinea As String
    
    Dim strURLTemp As String
    Dim blnEncontrado As Boolean
    
    On Error GoTo errores
    
    strURLTemp = Environ("tmp") & "\" & fso.GetTempName() & ".txt"
    strComando = "cmd /c whoami /user >" & Chr(34) & strURLTemp & Chr(34)
    EjecutarShell strComando, EnumSino.Sí
    Set fichero = fso.OpenTextFile(strURLTemp)
    Do Until fichero.AtEndOfStream
        strLinea = fichero.ReadLine
        If InStr(1, strLinea, "\") <> 0 Then
            blnEncontrado = True
            GoTo fin
        End If
    Loop
    blnEncontrado = False
fin:
    fichero.Close
    Set fichero = Nothing
    fso.DeleteFile strURLTemp, True
    
    If blnEncontrado = False Then
        getSSIDporMaquina = ""
        Exit Function
    End If
    If InStr(1, strLinea, " ") = 0 Then
        getSSIDporMaquina = ""
        Exit Function
    End If
    dato = Split(strLinea, " ")
    
    getSSIDporMaquina = Trim(dato(UBound(dato)))
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSSIDporMaquina ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Public Function getCorreoPorSSID(Optional p_SSID As String, Optional ByRef p_Error As String) As String

    Dim m_Correo As String
    
        
    On Error GoTo errores
    If p_SSID = "" Then
        getCorreoPorSSID = getCorreoPorSSIDPorArchivo(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    m_Correo = getCorreoPorSSIDEnBBDD(p_SSID, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Correo <> "" Then
        getCorreoPorSSID = m_Correo
        Exit Function
    End If
    getCorreoPorSSID = getCorreoPorSSIDPorArchivo(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCorreoPorSSID ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getCorreoPorSSIDPorArchivo(Optional ByRef p_Error As String) As String

    Dim dato As Variant
    Dim m_Texto As String
    Dim m_URL As String
    
    On Error GoTo errores
    If m_ObjEntorno Is Nothing Then
        Exit Function
    End If
    m_URL = m_ObjEntorno.URLArchivoSSID
    If Not fso.FileExists(m_URL) Then
        Exit Function
    End If
    
    m_Texto = getTextoArchivo(m_URL, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    dato = Split(m_Texto, "|")
    If UBound(dato) < 1 Then
        Exit Function
    End If
    getCorreoPorSSIDPorArchivo = dato(1)
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCorreoPorSSIDPorArchivo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getCorreoPorSSIDEnBBDD(p_SSID As String, Optional ByRef p_Error As String) As String

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
        
    On Error GoTo errores
    If p_SSID = "" Then
        Exit Function
    End If
    m_SQL = "SELECT correousuario " & _
            "FROM TbConexionesRegistro " & _
            "WHERE UsuarioSSID='" & p_SSID & "' " & _
            "ORDER BY FechaConexion DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        getCorreoPorSSIDEnBBDD = Nz(.Fields("correousuario"), "")
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCorreoPorSSIDEnBBDD ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getSSID(Optional p_URL As String, Optional ByRef p_Error As String) As String
    
    Dim dato As Variant
    Dim m_Texto As String
    
    
    On Error GoTo errores
    If p_URL = "" Then
        getSSID = getSSIDporMaquina
        Exit Function
    End If
    If Not fso.FileExists(p_URL) Then
        Exit Function
    End If
    m_Texto = getTextoArchivo(p_URL, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    dato = Split(m_Texto, "|")
    getSSID = dato(0)
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSSID ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function getFechaDeArchivoSSID(p_URL As String, Optional ByRef p_Error As String) As String
    
    Dim dato As Variant
    Dim m_Texto As String
    
    
    On Error GoTo errores
    
    If Not fso.FileExists(p_URL) Then
        Exit Function
    End If
    m_Texto = getTextoArchivo(p_URL, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    dato = Split(m_Texto, "|")
    If UBound(dato) < 2 Then
        Exit Function
    End If
    
    getFechaDeArchivoSSID = dato(2)
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getFechaDeArchivoSSID ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function


Public Function MismoSSID(Optional ByRef p_Error As String) As Boolean
    
    Dim m_URL As String
    Dim m_SSIDEquipo As String
    Dim m_SSIDArchivo As String
    
    On Error GoTo errores
    If m_ObjEntorno Is Nothing Then
        MismoSSID = False
        Exit Function
    End If
    m_URL = m_ObjEntorno.URLArchivoSSID
    If Not fso.FileExists(m_URL) Then
        MismoSSID = False
        Exit Function
    End If
    m_SSIDEquipo = getSSID(p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SSIDArchivo = getSSID(p_URL:=m_URL, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_SSIDEquipo = m_SSIDArchivo Then
        MismoSSID = True
    Else
        MismoSSID = False
    End If
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MismoSSID ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function MismoDia(Optional ByRef p_Error As String) As Boolean
    
    Dim m_Texto As String
    Dim m_Fecha As String
    Dim m_Dia As String
    Dim m_Mes As String
    Dim m_Año As String
    
    On Error GoTo errores
    
    m_Fecha = getFechaDeArchivoSSID(m_ObjEntorno.URLArchivoSSID, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not IsDate(m_Fecha) Then
        MismoDia = False
        Exit Function
    End If
    If Year(m_Fecha) <> Year(Date) Then
        MismoDia = False
        Exit Function
    End If
    If Month(m_Fecha) <> Month(Date) Then
        MismoDia = False
        Exit Function
    End If
    If Day(m_Fecha) <> Day(Date) Then
        MismoDia = False
        Exit Function
    End If
   MismoDia = True
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MismoDia ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function AperturaSinLogin(Optional ByRef p_Error As String) As Boolean
    
    Dim m_PermitirDiaSinContrasenia As Boolean
    Dim m_MismoSSID As Boolean
    Dim m_MismoDia As Boolean
   
    On Error GoTo errores
    If m_ObjEntorno.EstablecidoJSON = False Then
        AperturaSinLogin = False
        Exit Function
    End If
    m_PermitirDiaSinContrasenia = m_ObjEntorno.PermitirDiaSinContrasenia
    If m_PermitirDiaSinContrasenia = False Then
        AperturaSinLogin = False
        Exit Function
    End If
    m_MismoSSID = MismoSSID(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_MismoSSID = False Then
        AperturaSinLogin = False
        Exit Function
    End If
    
    m_MismoDia = MismoDia(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_MismoDia = False Then
        AperturaSinLogin = False
        Exit Function
    End If
   AperturaSinLogin = True
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AperturaSinLogin ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function EstablecidoJSON(Optional ByRef p_Error As String) As Boolean
    
    Dim m_URL As String
   
    On Error GoTo errores
    m_URL = m_ObjEntorno.URLDirAplicaciones & "configuracion.json"
    If Not fso.FileExists(m_URL) Then
        EstablecidoJSON = False
    Else
        EstablecidoJSON = True
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecidoJSON ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Public Function getTextoArchivo(p_URL As String, Optional ByRef p_Error As String) As String
    
    Dim archivo As Integer
    
    On Error GoTo errores
    If Not fso.FileExists(p_URL) Then
        p_Error = "No es alcanzable el archivo"
        Err.Raise 1000
    End If
    archivo = FreeFile()
    Open p_URL For Input As #archivo
    getTextoArchivo = Input$(LOF(archivo), archivo)
    Close #archivo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getTextoArchivo ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function getParametroConf(p_parametro As String, Optional ByRef p_Error As String)
    
    Dim m_URL As String
    Dim jsonTexto As String
    Dim jsonObject As Object
    
    On Error GoTo errores
    If p_parametro = "" Then
        Exit Function
    End If
    If m_ObjEntorno Is Nothing Then
        Exit Function
    End If
    jsonTexto = getTextoArchivo(m_ObjEntorno.URLConfiguracion, p_Error)
    If p_Error <> "" Then
        Exit Function
    End If
    If InStr(1, jsonTexto, "{") = 0 Then
        Exit Function
    End If
    Set jsonObject = JsonConverter.ParseJson(jsonTexto)
    getParametroConf = jsonObject("Lanzadera")(p_parametro)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getParametroConf ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Public Function SetSSOFAArchivo(p_Correo As String, Optional ByRef p_Error As String) As String
    
    
    Dim fichero As Scripting.TextStream
    Dim m_SSID As String
    Dim m_ArchivoSSID As String
    
    
    On Error GoTo errores
    If InStr(1, p_Correo, "@") = 0 Then
        Exit Function
    End If
    m_SSID = getSSID()
    
    If m_ObjEntorno Is Nothing Then
        Exit Function
    End If
    
    m_ArchivoSSID = m_ObjEntorno.URLArchivoSSID
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If fso.FileExists(m_ArchivoSSID) Then
        fso.DeleteFile m_ArchivoSSID
    End If
    Set fichero = fso.CreateTextFile(m_ArchivoSSID)
    
    fichero.WriteLine m_SSID & "|" & p_Correo & "|" & Date
    
    fichero.Close
    

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SetSSOFAArchivo ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function



Public Function EnOficina(Optional ByRef p_Error As String) As EnumSino
    
    Dim strIPS As String
    On Error GoTo errores
    strIPS = GetIPAddresses
    If InStr(1, strIPS, SubRedOficina) = 0 Then
        EnOficina = EnumSino.No
    Else
        EnOficina = EnumSino.Sí
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnOficina ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function getUsuarioRedConectado(Optional p_UsuarioDeRed As String, Optional ByRef p_Error As String) As String
    
    Dim objNetwork As Object
    On Error GoTo errores
    
    If p_UsuarioDeRed <> "" Then
        getUsuarioRedConectado = p_UsuarioDeRed
        Exit Function
    End If
    
    Set objNetwork = CreateObject("Wscript.Network")
    getUsuarioRedConectado = objNetwork.UserName
    Set objNetwork = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioRedConectado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function FormatoCorreoCorrecto( _
                                        ByVal p_email As String, _
                                        Optional ByRef p_Error As String _
                                        ) As EnumSino
    
    
    On Error GoTo errores
    If InStr(1, p_email, ".com_") <> 0 Then
        dato = Split(p_email, ".com_")
        p_email = dato(0) & ".com"
    End If
    If Right(p_email, 15) <> "@telefonica.com" Then
        FormatoCorreoCorrecto = EnumSino.No
    Else
        dato = Split(p_email, "@telefonica.com")
        If InStr(dato(0), ".") = 0 Then
            FormatoCorreoCorrecto = EnumSino.No
        Else
            FormatoCorreoCorrecto = EnumSino.Sí
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FormatoCorreoCorrecto ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function GenerarPassValida( _
                                    Cantidad As Integer, _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    Dim vecesRandom As Integer
    Dim i As Integer
    Dim Seg As Integer
    Dim Aleatorio As Double
    Dim divisor As String
    Dim intDigitosSeg As Integer
    Dim intSuperior As Integer
    Dim intInferior As Integer
    Dim m_Pass As String
    
    On Error GoTo errores
    
    Seg = CInt(Right(Round(Timer, 0), 2))
    Aleatorio = CDbl(Seg) / 100
    intSuperior = 100
    intInferior = 1
    vecesRandom = Int((intSuperior - intInferior + 1) * Aleatorio + intInferior)
    For i = 1 To vecesRandom
        m_Pass = GenerarPassAleatoria(Cantidad, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Next
    'Asignamos a la función los caractere concatenados.
    GenerarPassValida = m_Pass
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarPassValida ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Private Function GenerarPassAleatoria( _
                                        p_Cantidad As Integer, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_Caracter As String
    Dim i As Integer
    Dim m_PassAleatoria As String
    Dim veces As Integer
    Dim TramoEntreMayusculasOMinusculasONumeros As Integer
    Dim strPass As String
    Dim NumeroAleatorio As Integer
    Const ASCIIMayusculasMax As Integer = 90
    Const ASCIIMayusculasMin As Integer = 65
    Const ASCIIMinusculasMax As Integer = 122
    Const ASCIIMinusculasMin As Integer = 97
    Const ASCIINumerosMax As Integer = 57
    Const ASCIINumerosMin As Integer = 48
    On Error Resume Next
    veces = 0
    
inicio:
    m_PassAleatoria = ""
    'Recorremos de 1 a la p_Cantidad especificada en el parámetro p_Cantidad
    For i = 1 To p_Cantidad
        TramoEntreMayusculasOMinusculasONumeros = Int((3 * Rnd) + 1)
        'TramoEntreMayusculasOMinusculasONumeros=1--->Minusculas
        'TramoEntreMayusculasOMinusculasONumeros=2--->Mayúsculas
        'TramoEntreMayusculasOMinusculasONumeros=3--->Números
        If TramoEntreMayusculasOMinusculasONumeros = 1 Then
            NumeroAleatorio = Int((ASCIIMinusculasMax - ASCIIMinusculasMin + 1) * Rnd + ASCIIMinusculasMin)
        ElseIf TramoEntreMayusculasOMinusculasONumeros = 2 Then
            NumeroAleatorio = Int((ASCIIMayusculasMax - ASCIIMayusculasMin + 1) * Rnd + ASCIIMayusculasMin)
        Else
            NumeroAleatorio = Int((ASCIINumerosMax - ASCIINumerosMin + 1) * Rnd + ASCIINumerosMin)
        End If
        m_Caracter = Chr(NumeroAleatorio)
        If m_PassAleatoria = "" Then
            m_PassAleatoria = m_Caracter
        Else
            m_PassAleatoria = m_PassAleatoria & m_Caracter
        End If
    Next
    If PassFormatoValido(m_PassAleatoria, p_Error) <> EnumSino.Sí Then
        veces = veces + 1
        If veces <= 100 Then
            GoTo inicio
        End If
    End If
    'Asignamos a la función los caractere concatenados.
    GenerarPassAleatoria = m_PassAleatoria
'   GenerarPassValida = strPass
   Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarPassAleatoria ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function PassFormatoValido( _
                                    p_pass As String, _
                                    Optional ByRef p_Error As String _
                                    ) As EnumSino
    
    Dim RegEx As Object
    Set RegEx = CreateObject("vbscript.regexp")
    On Error GoTo errores
    With RegEx
        .Global = True
        .Pattern = "^(?=\w*\d)(?=\w*[A-Z])(?=\w*[a-z])\S{8,16}$"
    End With
    If RegEx.Test(p_pass) = False Then
        PassFormatoValido = EnumSino.No
    Else
        PassFormatoValido = EnumSino.Sí
    End If
    Set RegEx = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PassFormatoValido ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function






Public Function DameHTMLParaCorreo( _
                                    p_Texto As String, _
                                    Optional p_Titulo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_mensaje As String
    On Error GoTo errores
     m_mensaje = "<!DOCTYPE html>" & vbNewLine
    m_mensaje = m_mensaje & "<html lang=""es"">" & vbNewLine
    m_mensaje = m_mensaje & "<head>" & vbNewLine
        m_mensaje = m_mensaje & "<title>" & p_Titulo & "</title>" & vbNewLine
        m_mensaje = m_mensaje & "<meta charset=""ISO-8859-1"" />" & vbNewLine
       
    m_mensaje = m_mensaje & "</head>" & vbNewLine
    m_mensaje = m_mensaje & "<body>" & vbNewLine
    m_mensaje = m_mensaje & "<pre> <STRONG>" & p_Texto & "</STRONG></pre>"
    
    DameHTMLParaCorreo = m_mensaje
     Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameHTMLParaCorreo ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function



Public Function DameHTMLCabecera( _
                                Optional p_Titulo As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim m_mensaje As String
   
    
    On Error GoTo errores
     m_mensaje = "<!DOCTYPE html>" & vbNewLine
    m_mensaje = m_mensaje & "<html lang=""es"">" & vbNewLine
    m_mensaje = m_mensaje & "<head>" & vbNewLine
    m_mensaje = m_mensaje & "<meta charset=""ISO-8859-1"">" & vbNewLine
    m_mensaje = m_mensaje & "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" & vbNewLine
    m_mensaje = m_mensaje & "<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbNewLine
    m_mensaje = m_mensaje & "<title>" & p_Titulo & "</title>" & vbNewLine
    m_mensaje = m_mensaje & "<style type=""text/css"">" & vbNewLine
        m_mensaje = m_mensaje & m_ObjEntorno.CSS & vbNewLine
    m_mensaje = m_mensaje & "</style>" & vbNewLine
    m_mensaje = m_mensaje & "</head>" & vbNewLine
    m_mensaje = m_mensaje & "<body>" & vbNewLine
    
    DameHTMLCabecera = m_mensaje
     Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameHTMLCabecera ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function DameFinHTML( _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim m_mensaje As String
   
    
    On Error GoTo errores
    m_mensaje = m_mensaje & "<p><strong>¡¡¡NO RESPONDA A ESTE CORREO (es un mensaje automático)!!!</strong></p>" & vbNewLine
    m_mensaje = m_mensaje & "</body>" & vbNewLine
    m_mensaje = m_mensaje & "</html>" & vbNewLine
    
    DameFinHTML = m_mensaje
     Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameFinHTML ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function DameHTMLUsuarioConectado( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim m_mensaje As String
   
    
    On Error GoTo errores
    m_mensaje = "<a href='mailto:" & m_ObjUsuarioConectadoLogin.CorreoUsuario & _
                "'>correo enviado por LANZADERA en nombre de: " & m_ObjUsuarioConectadoLogin.Nombre & "</a>" & vbNewLine
    DameHTMLUsuarioConectado = m_mensaje
     Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameHTMLUsuarioConectado ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function CorreoCuentaBloqueada( _
                                        p_ObjUsuario As usuario, _
                                        Optional ByRef p_Error As String _
                                        ) As String
        
    Dim m_HTMLCabecera As String
    
    Dim m_HTMLTexto As String
    Dim m_HTMLFin As String
    Dim m_mensaje As String
    Dim m_Destinatarios As String
    Dim m_Asunto As String
    Dim m_ObjCorreo As Correo
    On Error GoTo errores
    
    m_Asunto = "Usuario Bloqueado por exceso de intentos de acceso (LANZADERA)"
    m_Destinatarios = "andres.romandelperal@telefonica.com"
    m_HTMLCabecera = DameHTMLCabecera("Cuenta bloqueada de Usuario", p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTexto = "El usuario: " & p_ObjUsuario.CorreoUsuario & " ha bloqueado su cuenta por exceso de intentos"
    m_HTMLFin = DameFinHTML(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_mensaje = m_HTMLCabecera & vbNewLine
    m_mensaje = m_mensaje & m_HTMLTexto & vbNewLine
    m_mensaje = m_mensaje & m_HTMLFin & vbNewLine
    
    Set m_ObjCorreo = New Correo
    With m_ObjCorreo
        .Asunto = m_Asunto
        .Cuerpo = m_mensaje
        .Destinatarios = m_Destinatarios
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set m_ObjCorreo = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoCuentaBloqueada ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function Login( _
                        ByVal p_ObjUsuario As usuario, _
                        Optional ByRef p_PassPlana As String, _
                        Optional ByRef p_Error As String _
                        ) As EnumPassEstados
    
    Dim m_PassEncriptada As String
    Dim m_PassCaducada As EnumSino
    Dim m_TieneQueCambiarLaPass As EnumSino
    Dim m_FechaBloqueo As String
    On Error GoTo errores
    
    If p_ObjUsuario Is Nothing Then
        p_Error = "Se ha de indicar el usuario"
        Err.Raise 1000
    End If
    If p_PassPlana <> "" Then
       m_PassEncriptada = SHA256(p_PassPlana)
        If m_PassEncriptada = "" Then
            p_Error = "No se ha podido encriptar la Pass"
            Err.Raise 1000
        End If
        If p_ObjUsuario.Password <> "" Then
            If p_ObjUsuario.Password <> m_PassEncriptada Then
                Login = EnumPassEstados.Incorrecto
                If m_ObjEntorno.NumeroLogins > 4 Then
                    p_ObjUsuario.Bloquear p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    m_ObjEntorno.NumeroLogins = 0
                Else
                    m_ObjEntorno.NumeroLogins = m_ObjEntorno.NumeroLogins + 1
                End If
                Exit Function
            End If
            
        End If
        m_FechaBloqueo = p_ObjUsuario.FechaBloqueo
        If IsDate(m_FechaBloqueo) Then
            If DateDiff("m", Now(), CDate(m_FechaBloqueo)) < 20 Then
                Login = EnumPassEstados.Bloqueada
                Exit Function
            End If
            p_ObjUsuario.DesBloquear p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_ObjEntorno.NumeroLogins = 0
        End If
    End If
    
    
   
    
    m_TieneQueCambiarLaPass = p_ObjUsuario.TieneQueCambiarLaPassCalculado
    p_Error = p_ObjUsuario.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TieneQueCambiarLaPass = 0 Then
        p_Error = "No se ha podido determinar si tiene una Pass de un sólo uso"
        Err.Raise 1000
    End If
    If m_TieneQueCambiarLaPass <> EnumSino.Sí Then
        Login = EnumPassEstados.Correcto
        Exit Function
    End If
    If m_TieneQueCambiarLaPass = EnumSino.Sí Then
        If p_ObjUsuario.PassIncialPlana = "" Then
            p_Error = "No se ha podido determinar si tiene una Pass de un sólo uso"
            Err.Raise 1000
        End If
        If p_ObjUsuario.PassIncialPlana <> p_PassPlana Then
            Login = EnumPassEstados.Incorrecto
        Else
            Login = EnumPassEstados.NecesitaCambio
        End If
        Exit Function
    End If
     m_PassCaducada = p_ObjUsuario.PassCaducada
    p_Error = p_ObjUsuario.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_PassCaducada = 0 Then
        p_Error = "No se ha podido determinar si la cuenta está caducada"
        Err.Raise 1000
    End If
    If m_PassCaducada = EnumSino.Sí Then
        Login = EnumPassEstados.caducada
        Exit Function
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Login ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Public Function CorreoAlAdministrador( _
                                        m_MensajeError As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
   
    Dim m_mensaje As String
    Dim m_Asunto As String
    Dim m_Nombre As String
    Dim m_NombreFormulario As String
    Dim m_TextoEnOficina As String
   
    
    Dim m_ObjCorreo As Correo
    On Error GoTo errores
    
    If m_MensajeError = "" Then
        p_Error = "No hay mensaje que enviar"
        Err.Raise 1000
    End If
    
    If m_EnOficina = Empty Then
        m_EnOficina = EnOficina(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If m_EnOficina = Empty Then
        m_TextoEnOficina = "En Oficina Desconocido"
    Else
        If m_EnOficina = EnumSino.Sí Then
            m_TextoEnOficina = "En Oficina"
        Else
            m_TextoEnOficina = "Fuera de Oficina"
        End If
    End If
    On Error Resume Next
    m_NombreFormulario = Application.Screen.ActiveForm.Name
    If Err.Number <> 0 Then
        Err.Clear
        m_NombreFormulario = "Desconocido"
    End If
    On Error GoTo errores
    m_Nombre = getNombreUsuarioConectado()
    
    m_Asunto = "Error en LANZADERA " & m_Nombre & " " & m_TextoEnOficina
    m_mensaje = "FORMULARIO del ERROR: " & m_NombreFormulario & vbNewLine
    m_mensaje = m_mensaje & "<BR> </BR>" & vbNewLine
    On Error Resume Next
    m_mensaje = m_mensaje & "NOMBRE EQUIPO: " & VBA.Environ("COMPUTERNAME")
    m_mensaje = m_mensaje & "<BR> </BR>" & vbNewLine
    On Error GoTo errores
    
    m_mensaje = m_mensaje & "DETALLE: " & m_MensajeError
    
    Set m_ObjCorreo = New Correo
    With m_ObjCorreo
        .Asunto = m_Asunto
        .Cuerpo = m_mensaje
        .Destinatarios = "ardelperal@gmail.com;andres.romandelperal@telefonica.com"
        .FechaGrabacion = Now()
        .EnviarCorreo p_Error
    End With
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoAlAdministrador ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function CorreoNoAccesoAVideos( _
                                        m_MensajeError As String, _
                                        m_NombreUsuario As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
   
    Dim m_mensaje As String
    Dim m_Asunto As String
    Dim m_ObjCorreo As Correo
    
    On Error GoTo errores
    
    If m_MensajeError = "" Then
        p_Error = "No hay mensaje que enviar"
        Err.Raise 1000
    End If
    
    
    m_Asunto = "Accesibilidad a Videos: " & m_NombreUsuario
    m_Asunto = m_Asunto & " (LANZADERA)"
    
    m_mensaje = m_MensajeError
    
    
    Set m_ObjCorreo = New Correo
    With m_ObjCorreo
        .Asunto = m_Asunto
        .Cuerpo = m_mensaje
        .Destinatarios = "andres.romanDELperal@telefonica.com"
        .FechaGrabacion = Now()
        .EnviarCorreo p_Error
    End With
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoNoAccesoAVideos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getDatoAleatorio( _
                                    p_EnumObjetos As EnumObjetos, _
                                    Optional ByRef p_Error _
                                    ) As String
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    Dim m_FilaInicial As Long
    Dim m_FilaFinal As Long
    Dim m_FilaAleatoria As Long
    Dim m_Fila As Long
    Dim m_NombreTabla As String
    Dim m_NombreCampo As String
    On Error GoTo errores
     p_Error = ""
     
    If p_EnumObjetos = EnumObjetos.Aplicacion Then
        m_NombreTabla = "TbAplicaciones"
        m_NombreCampo = "IDAplicacion"
    ElseIf p_EnumObjetos = EnumObjetos.Conexion Then
        m_NombreTabla = "TbConexionesRegistro"
        m_NombreCampo = "IDConexion"
    ElseIf p_EnumObjetos = EnumObjetos.Correo Then
        m_NombreTabla = "TbUsuariosCorreosEnvio"
        m_NombreCampo = "IDCorreo"
    ElseIf p_EnumObjetos = EnumObjetos.usuario Then
        m_NombreTabla = "TbUsuariosAplicaciones"
        m_NombreCampo = "CorreoUsuario"
    Else
        p_Error = "Objeto fuera del test"
        Err.Raise 1000
    End If

    m_SQL = "SELECT " & m_NombreTabla & "." & m_NombreCampo & " " & _
            "FROM " & m_NombreTabla & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveLast
        .MoveFirst
        m_FilaFinal = .RecordCount
        m_FilaInicial = 1
        m_FilaAleatoria = Int((m_FilaFinal - m_FilaInicial + 1) * Rnd + m_FilaInicial)
        m_Fila = 1
        Do While Not .EOF
            If .AbsolutePosition + 1 = m_FilaAleatoria Then
                getDatoAleatorio = Nz(.Fields(m_NombreCampo), "")
                rcdDatos.Close
                Set rcdDatos = Nothing
                Exit Function
            End If
        
            .MoveNext
            
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDatoAleatorio ha devuelto un error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function Dame( _
                        p_NOmbreTabla As String, _
                        p_NombreCampo As String, _
                        p_NombreCampoID As String, _
                        p_ValorCampoID As String, _
                        Optional ByRef p_Error _
                        ) As String
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    
    On Error Resume Next
                        
    m_SQL = "SELECT " & p_NOmbreTabla & "." & p_NombreCampo & " " & _
            "FROM " & p_NOmbreTabla & " " & _
            "WHERE " & p_NOmbreTabla & "." & p_NombreCampoID & "=" & p_ValorCampoID & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If Err.Number <> 0 Then
        Err.Clear
        m_SQL = "SELECT " & p_NOmbreTabla & "." & p_NombreCampo & " " & _
                "FROM " & p_NOmbreTabla & " " & _
                "WHERE " & p_NOmbreTabla & "." & p_NombreCampoID & "='" & p_ValorCampoID & "';"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        If Err.Number <> 0 Then
            GoTo errores
        End If
        If Not rcdDatos.EOF Then
            Dame = Nz(rcdDatos.Fields(p_NombreCampo), "")
        End If
    Else
        If Not rcdDatos.EOF Then
            Dame = Nz(rcdDatos.Fields(p_NombreCampo), "")
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Dame ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function


Public Function UsuarioAltaEdicionEstado(Optional ByRef p_Error As String) As String
    
    Dim m_FormularioGeneral As Form
    Dim m_FormConfiguracion As Form
    
    
    On Error GoTo errores
    
    Set m_FormularioGeneral = Forms("Form00Inicial")
    If m_FormularioGeneral.Controls("FormDetalle").SourceObject = "Form00ConfiguracionAdmin00" Then
        Set m_FormConfiguracion = m_FormularioGeneral.Controls("FormDetalle").Form
        If m_FormConfiguracion.Controls("CuadroSuperior").SourceObject <> "" Then
            m_FormConfiguracion.Controls("CuadroSuperior").SourceObject = ""
        End If
        m_FormConfiguracion.Controls("CuadroSuperior").SourceObject = "FormNombreUsuario"
        If m_FormConfiguracion.Controls("CuadroMedio").SourceObject <> "" Then
            m_FormConfiguracion.Controls("CuadroMedio").SourceObject = ""
        End If
        m_FormConfiguracion.Controls("CuadroMedio").SourceObject = "FormUsuario"
        If m_FormConfiguracion.Controls("CuadroInferior").SourceObject <> "" Then
            m_FormConfiguracion.Controls("CuadroInferior").SourceObject = ""
        End If
        m_FormConfiguracion.Controls("CuadroInferior").SourceObject = "FormUsuarioPerfilAplicaciones"
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioAltaEdicionEstado ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function

Public Function UsuarioBorradoEstado(Optional ByRef p_Error As String) As String
    
    Dim m_FormularioGeneral As Form
    Dim m_FormConfiguracion As Form
    
    On Error GoTo errores
    
    Set m_FormularioGeneral = Forms("Form00Inicial")
    If m_FormularioGeneral.Controls("FormDetalle").SourceObject = "Form00ConfiguracionAdmin00" Then
        Set m_FormConfiguracion = m_FormularioGeneral.Controls("FormDetalle").Form
        Set m_FormConfiguracion = m_FormConfiguracion.Controls("FormDetalle").Form
        If m_FormConfiguracion.Controls("CuadroSuperior").SourceObject <> "" Then
            m_FormConfiguracion.Controls("CuadroSuperior").SourceObject = ""
        End If
        m_FormConfiguracion.Controls("CuadroSuperior").SourceObject = "FormNombreUsuario"
        If m_FormConfiguracion.Controls("CuadroMedio").SourceObject <> "" Then
            m_FormConfiguracion.Controls("CuadroMedio").SourceObject = ""
        End If
        If m_FormConfiguracion.Controls("CuadroInferior").SourceObject <> "" Then
            m_FormConfiguracion.Controls("CuadroInferior").SourceObject = ""
        End If
       
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioBorradoEstado ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function

Public Function UsuarioSeleccionadoEstado(Optional ByRef p_Error As String) As String
    
    Dim m_FormularioGeneral As Form
    Dim m_FormConfiguracion As Form
    Dim m_FormNombreUsuario As Form
    
    On Error GoTo errores
    If m_ObjUsuarioActivo Is Nothing Then
        p_Error = "No se sabe qué usuario ha sido seleccionado"
        Err.Raise 1000
    End If
    Set m_FormularioGeneral = Forms("Form00Inicial")
    If m_FormularioGeneral.Controls("FormDetalle").SourceObject = "Form00ConfiguracionAdmin00" Then
        Set m_FormConfiguracion = m_FormularioGeneral.Controls("FormDetalle").Form
        If m_FormConfiguracion.Controls("CuadroSuperior").SourceObject <> "" Then
            m_FormConfiguracion.Controls("CuadroSuperior").SourceObject = ""
        End If
        m_FormConfiguracion.Controls("CuadroSuperior").SourceObject = "FormNombreUsuario"
        Set m_FormNombreUsuario = m_FormConfiguracion.Controls("CuadroSuperior").Form
        m_FormNombreUsuario.Controls("Nombre") = Null
        m_FormNombreUsuario.Controls("Usuario") = m_ObjUsuarioActivo.Nombre
        m_FormNombreUsuario.Controls("ComandoRestablecerPassAUsuario").Enabled = True
        m_FormNombreUsuario.Controls("ComandoEnviarCorreo").Enabled = True
        If m_ObjUsuarioActivo.UsuarioImborrable = True Then
            m_FormNombreUsuario.Controls("ComandoEliminar").Enabled = False
        Else
            m_FormNombreUsuario.Controls("ComandoEliminar").Enabled = True
        End If
        
        If m_FormConfiguracion.Controls("CuadroMedio").SourceObject <> "" Then
            m_FormConfiguracion.Controls("CuadroMedio").SourceObject = ""
        End If
        m_FormConfiguracion.Controls("CuadroMedio").SourceObject = "FormUsuario"
        If m_FormConfiguracion.Controls("CuadroInferior").SourceObject <> "" Then
            m_FormConfiguracion.Controls("CuadroInferior").SourceObject = ""
        End If
        m_FormConfiguracion.Controls("CuadroInferior").SourceObject = "FormUsuarioPerfilAplicaciones"
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioSeleccionadoEstado ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function

Public Function Lanzar( _
                        p_ObjBoton As CommandButton, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim m_objAplicacion As Aplicacion
    
    On Error GoTo errores
    
    Application.Screen.ActiveForm.AllowEdits = False
    'Me.AllowEdits = False
    If Not IsNumeric(Nz(p_ObjBoton.Tag, "")) Then
        p_Error = "No se puede obtener el ID de la aplicación"
        Err.Raise 1000
    End If
    
    Set m_objAplicacion = Constructor.getAplicacion(p_IDAplicacion:=CStr(p_ObjBoton.Tag), p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objAplicacion Is Nothing Then
        p_Error = "No se ha encontrado una aplicación con el ID: " & p_ObjBoton.Tag
        Err.Raise 1000
    End If
    
    If m_objAplicacion.EjecucionEnOficinaCalculado = EnumSino.Sí Then
        If m_EnOficina <> EnumSino.Sí Then
            p_Error = "Esta aplicación sólo se puede ejecutar desde la Oficina"
            Err.Raise 1000
        End If
    End If
    m_objAplicacion.Lanzar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Application.Screen.ActiveForm.AllowEdits = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Lanzar ha producido el error nº: " & Err.Number & _
                    vbCrLf & "Detalle: " & Err.Description
    End If
    Application.Screen.ActiveForm.AllowEdits = True
    Pregunta = MsgBox(p_Error, vbCritical, "Error")

End Function


Public Function LanzarGestionRiesgosPrueba( _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_URLIniLocal As String
    Dim m_URLIniRemoto As String
    
    Dim m_URLIconoLocal As String
    Dim m_URLIconoRemoto As String
    
    Dim m_URLbmpLocal As String
    Dim m_URLbmpRemoto As String
    
    Dim m_URLEjecutableRemoto As String
    Dim m_URLEjecutableLocal As String
    
    Dim m_URLDirectorioAplicacionUsuario As String
    Dim m_URLDirectorioRemoto  As String
    
    Dim m_HayQueCopiarALocal As EnumSino
    Dim m_db As DAO.Database
    Dim wks As DAO.Workspace
    Dim m_ComandoEjecutable As String
    Dim m_VersionLocal As String
    Dim m_VersionRemota As String
    Dim m_Pass As String
    
    On Error GoTo errores
    m_Pass = GetBackendPassword(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Application.Screen.ActiveForm.AllowEdits = False
   
    If m_CorreoUsuarioEnLogin = "" Then
        If m_ObjUsuarioConectadoLogin Is Nothing Then
            p_Error = "No se sabe el usuario que ha hecho login"
            Err.Raise 1000
        End If
        m_CorreoUsuarioEnLogin = m_ObjUsuarioConectadoLogin.CorreoUsuario
    End If
    If InStr(1, m_CorreoUsuarioEnLogin, "@") = 0 Then
        p_Error = "No se sabe el usuario que ha hecho login"
        Err.Raise 1000
    End If
    m_URLDirectorioAplicacionUsuario = Environ("APPDATA") & "\" & "Aplicaciones DYSN\GESTION RIESGOS PRUEBA\"
    If Not fso.FolderExists(m_URLDirectorioAplicacionUsuario) Then
        fso.CreateFolder m_URLDirectorioAplicacionUsuario
    End If
    m_URLDirectorioRemoto = NormalizarRuta(m_URLRutaAplicacionesRemotas) & "GESTION RIESGOS PRUEBA\recursos\"
    m_URLIniRemoto = m_URLDirectorioRemoto & "Gestion_Riesgos.ini"
    If Not fso.FileExists(m_URLIniRemoto) Then
        p_Error = "No se sabe la versión del servidor"
        Err.Raise 1000
    End If
    m_URLEjecutableRemoto = m_URLDirectorioRemoto & "Gestion_Riesgos.accde"
   
    If Not fso.FileExists(m_URLEjecutableRemoto) Then
        p_Error = "No se localiza el ejecutable en el servidor"
        Err.Raise 1000
    End If
    m_URLIconoRemoto = m_URLDirectorioRemoto & "IconoAplicacion.ico"
    If Not fso.FileExists(m_URLIniRemoto) Then
        p_Error = "No se puede obtener el icono remoto"
        Err.Raise 1000
    End If
    m_URLbmpRemoto = m_URLDirectorioRemoto & "Gestion_Riesgos.bmp"
    If Not fso.FileExists(m_URLIniRemoto) Then
        p_Error = "No se puede obtener el icono remoto"
        Err.Raise 1000
    End If
    m_URLIniLocal = m_URLDirectorioAplicacionUsuario & "Gestion_Riesgos.ini"
    m_URLbmpLocal = m_URLDirectorioAplicacionUsuario & "Gestion_Riesgos.bmp"
    m_URLEjecutableLocal = m_URLDirectorioAplicacionUsuario & "Gestion_Riesgos.accde"
    m_URLIconoLocal = m_URLDirectorioAplicacionUsuario & "IconoAplicacion.ico"
    
    m_HayQueCopiarALocal = EnumSino.No
    m_VersionLocal = getVersion(m_URLIniLocal, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_VersionLocal = "" Then
        GoTo CopiarALocal
    End If
    m_VersionRemota = getVersion(m_URLIniRemoto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_VersionLocal <> m_VersionRemota Then
        GoTo CopiarALocal
    End If
    If Not fso.FileExists(m_URLbmpLocal) Or Not fso.FileExists(m_URLEjecutableLocal) Or _
            Not fso.FileExists(m_URLIconoLocal) Then
        GoTo CopiarALocal
    End If
    GoTo EjecutarLocal
   
    
CopiarALocal:
    m_HayQueCopiarALocal = EnumSino.Sí
    fso.CopyFile m_URLEjecutableRemoto, m_URLEjecutableLocal, True
    fso.CopyFile m_URLIconoRemoto, m_URLIconoLocal, True
    fso.CopyFile m_URLbmpRemoto, m_URLbmpLocal, True
    fso.CopyFile m_URLIniRemoto, m_URLIniLocal, True
    Set wks = DBEngine.Workspaces(0)
    Set m_db = wks.OpenDatabase(m_URLEjecutableLocal, True, False, "MS Access;PWD=" & m_Pass)
    AlterarPropriedad m_URLEjecutableLocal, "AppIcon", dbText, m_URLIconoRemoto, m_db, m_Pass, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_db.NewPassword m_Pass, ""
    m_db.Close
    Set m_db = Nothing
    wks.Close
    Set wks = Nothing
EjecutarLocal:
    
 
    m_ComandoEjecutable = Chr(34) & m_ObjEntorno.URLEjecutableAccess & Chr(34) & " "
    m_ComandoEjecutable = m_ComandoEjecutable & Chr(34) & m_URLEjecutableLocal & Chr(34) & " " & " /cmd " & m_CorreoUsuarioEnLogin
    EjecutarShelllanzar m_ComandoEjecutable, EnumSino.No, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Application.Screen.ActiveForm.AllowEdits = True
    If m_ObjUsuarioConectadoLogin.MantenerLanzaderaAbierta = False Then
        
        DoCmd.Quit acQuitSaveAll
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método LanzarGestionRiesgosPrueba ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Application.Screen.ActiveForm.AllowEdits = True
    On Error Resume Next
    If Not m_db Is Nothing Then
        m_db.Close
        Set m_db = Nothing
    End If
    If Not wks Is Nothing Then
        wks.Close
        Set wks = Nothing
    End If
    
End Function

Public Function LanzarNCPrueba( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_URLIniLocal As String
    Dim m_URLIniRemoto As String
    
    Dim m_URLIconoLocal As String
    Dim m_URLIconoRemoto As String
    
    Dim m_URLbmpLocal As String
    Dim m_URLbmpRemoto As String
    
    Dim m_URLEjecutableRemoto As String
    Dim m_URLEjecutableLocal As String
    
    Dim m_URLDirectorioAplicacionUsuario As String
    Dim m_URLDirectorioRemoto  As String
    
    Dim m_HayQueCopiarALocal As EnumSino
    Dim m_db As DAO.Database
    Dim wks As DAO.Workspace
    Dim m_ComandoEjecutable As String
    Dim m_VersionLocal As String
    Dim m_VersionRemota As String
    Dim m_Pass As String
    
    On Error GoTo errores
    m_Pass = GetBackendPassword(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Application.Screen.ActiveForm.AllowEdits = False
   
    If m_CorreoUsuarioEnLogin = "" Then
        If m_ObjUsuarioConectadoLogin Is Nothing Then
            p_Error = "No se sabe el usuario que ha hecho login"
            Err.Raise 1000
        End If
        m_CorreoUsuarioEnLogin = m_ObjUsuarioConectadoLogin.CorreoUsuario
    End If
    If InStr(1, m_CorreoUsuarioEnLogin, "@") = 0 Then
        p_Error = "No se sabe el usuario que ha hecho login"
        Err.Raise 1000
    End If
    m_URLDirectorioAplicacionUsuario = Environ("APPDATA") & "\" & "Aplicaciones DYSN\No Conformidades PRUEBA\"
    If Not fso.FolderExists(m_URLDirectorioAplicacionUsuario) Then
        fso.CreateFolder m_URLDirectorioAplicacionUsuario
    End If
    m_URLDirectorioRemoto = NormalizarRuta(m_URLRutaAplicacionesRemotas) & "No Conformidades PRUEBA\recursos\"
    m_URLIniRemoto = m_URLDirectorioRemoto & "NoConformidades.ini"
    If Not fso.FileExists(m_URLIniRemoto) Then
        p_Error = "No se sabe la versión del servidor"
        Err.Raise 1000
    End If
    m_URLEjecutableRemoto = m_URLDirectorioRemoto & "NoConformidades.accde"
   
    If Not fso.FileExists(m_URLEjecutableRemoto) Then
        p_Error = "No se localiza el ejecutable en el servidor"
        Err.Raise 1000
    End If
    m_URLIconoRemoto = m_URLDirectorioRemoto & "IconoAplicacion.ico"
    If Not fso.FileExists(m_URLIniRemoto) Then
        p_Error = "No se puede obtener el icono remoto"
        Err.Raise 1000
    End If
    m_URLbmpRemoto = m_URLDirectorioRemoto & "NoConformidades.bmp"
    If Not fso.FileExists(m_URLIniRemoto) Then
        p_Error = "No se puede obtener el icono remoto"
        Err.Raise 1000
    End If
    m_URLIniLocal = m_URLDirectorioAplicacionUsuario & "NoConformidades.ini"
    m_URLbmpLocal = m_URLDirectorioAplicacionUsuario & "NoConformidades.bmp"
    m_URLEjecutableLocal = m_URLDirectorioAplicacionUsuario & "NoConformidades.accde"
    m_URLIconoLocal = m_URLDirectorioAplicacionUsuario & "IconoAplicacion.ico"
    
    m_HayQueCopiarALocal = EnumSino.No
    m_VersionLocal = getVersion(m_URLIniLocal, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_VersionLocal = "" Then
        GoTo CopiarALocal
    End If
    m_VersionRemota = getVersion(m_URLIniRemoto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_VersionLocal <> m_VersionRemota Then
        GoTo CopiarALocal
    End If
    If Not fso.FileExists(m_URLbmpLocal) Or Not fso.FileExists(m_URLEjecutableLocal) Or _
            Not fso.FileExists(m_URLIconoLocal) Then
        GoTo CopiarALocal
    End If
    GoTo EjecutarLocal
   
    
CopiarALocal:
    m_HayQueCopiarALocal = EnumSino.Sí
    fso.CopyFile m_URLEjecutableRemoto, m_URLEjecutableLocal, True
    fso.CopyFile m_URLIconoRemoto, m_URLIconoLocal, True
    fso.CopyFile m_URLbmpRemoto, m_URLbmpLocal, True
    fso.CopyFile m_URLIniRemoto, m_URLIniLocal, True
    Set wks = DBEngine.Workspaces(0)
    Set m_db = wks.OpenDatabase(m_URLEjecutableLocal, True, False, "MS Access;PWD=" & m_Pass)
    AlterarPropriedad m_URLEjecutableLocal, "AppIcon", dbText, m_URLIconoRemoto, m_db, m_Pass, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_db.NewPassword m_Pass, ""
    m_db.Close
    Set m_db = Nothing
    wks.Close
    Set wks = Nothing
EjecutarLocal:
    
 
    m_ComandoEjecutable = Chr(34) & m_ObjEntorno.URLEjecutableAccess & Chr(34) & " "
    m_ComandoEjecutable = m_ComandoEjecutable & Chr(34) & m_URLEjecutableLocal & Chr(34) & " " & " /cmd " & m_CorreoUsuarioEnLogin
    EjecutarShelllanzar m_ComandoEjecutable, EnumSino.No, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Application.Screen.ActiveForm.AllowEdits = True
    If m_ObjUsuarioConectadoLogin.MantenerLanzaderaAbierta = False Then
        
        DoCmd.Quit acQuitSaveAll
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método LanzarNCPrueba ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Application.Screen.ActiveForm.AllowEdits = True
    On Error Resume Next
    If Not m_db Is Nothing Then
        m_db.Close
        Set m_db = Nothing
    End If
    If Not wks Is Nothing Then
        wks.Close
        Set wks = Nothing
    End If
    
End Function

Public Function LanzarCondorPrueba( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_URLIniLocal As String
    Dim m_URLIniRemoto As String
    
    Dim m_URLIconoLocal As String
    Dim m_URLIconoRemoto As String
    
    Dim m_URLbmpLocal As String
    Dim m_URLbmpRemoto As String
    
    Dim m_URLEjecutableRemoto As String
    Dim m_URLEjecutableLocal As String
    
    Dim m_URLDirectorioAplicacionUsuario As String
    Dim m_URLDirectorioRemoto  As String
    
    Dim m_HayQueCopiarALocal As EnumSino
    Dim m_db As DAO.Database
    Dim wks As DAO.Workspace
    Dim m_ComandoEjecutable As String
    Dim m_VersionLocal As String
    Dim m_VersionRemota As String
    Dim m_Pass As String
    
    On Error GoTo errores
    m_Pass = GetBackendPassword(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    p_Error = "Aún no hay un producto mínimo viable"
    Err.Raise 1000
    Application.Screen.ActiveForm.AllowEdits = False
   
    If m_CorreoUsuarioEnLogin = "" Then
        If m_ObjUsuarioConectadoLogin Is Nothing Then
            p_Error = "No se sabe el usuario que ha hecho login"
            Err.Raise 1000
        End If
        m_CorreoUsuarioEnLogin = m_ObjUsuarioConectadoLogin.CorreoUsuario
    End If
    If InStr(1, m_CorreoUsuarioEnLogin, "@") = 0 Then
        p_Error = "No se sabe el usuario que ha hecho login"
        Err.Raise 1000
    End If
    m_URLDirectorioAplicacionUsuario = Environ("APPDATA") & "\" & "Aplicaciones DYSN\Concesiones PRUEBA\"
    If Not fso.FolderExists(m_URLDirectorioAplicacionUsuario) Then
        fso.CreateFolder m_URLDirectorioAplicacionUsuario
    End If
    m_URLDirectorioRemoto = NormalizarRuta(m_URLRutaAplicacionesRemotas) & "Concesiones PRUEBA\recursos\"
    m_URLIniRemoto = m_URLDirectorioRemoto & "Concesiones.ini"
    If Not fso.FileExists(m_URLIniRemoto) Then
        p_Error = "No se sabe la versión del servidor"
        Err.Raise 1000
    End If
    m_URLEjecutableRemoto = m_URLDirectorioRemoto & "Concesiones.accde"
   
    If Not fso.FileExists(m_URLEjecutableRemoto) Then
        p_Error = "No se localiza el ejecutable en el servidor"
        Err.Raise 1000
    End If
    m_URLIconoRemoto = m_URLDirectorioRemoto & "IconoAplicacion.ico"
    If Not fso.FileExists(m_URLIniRemoto) Then
        p_Error = "No se puede obtener el icono remoto"
        Err.Raise 1000
    End If
    m_URLbmpRemoto = m_URLDirectorioRemoto & "Concesiones.bmp"
    If Not fso.FileExists(m_URLIniRemoto) Then
        p_Error = "No se puede obtener el icono remoto"
        Err.Raise 1000
    End If
    m_URLIniLocal = m_URLDirectorioAplicacionUsuario & "Concesiones.ini"
    m_URLbmpLocal = m_URLDirectorioAplicacionUsuario & "Concesiones.bmp"
    m_URLEjecutableLocal = m_URLDirectorioAplicacionUsuario & "Concesiones.accde"
    m_URLIconoLocal = m_URLDirectorioAplicacionUsuario & "IconoAplicacion.ico"
    
    m_HayQueCopiarALocal = EnumSino.No
    m_VersionLocal = getVersion(m_URLIniLocal, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_VersionLocal = "" Then
        GoTo CopiarALocal
    End If
    m_VersionRemota = getVersion(m_URLIniRemoto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_VersionLocal <> m_VersionRemota Then
        GoTo CopiarALocal
    End If
    If Not fso.FileExists(m_URLbmpLocal) Or Not fso.FileExists(m_URLEjecutableLocal) Or _
            Not fso.FileExists(m_URLIconoLocal) Then
        GoTo CopiarALocal
    End If
    GoTo EjecutarLocal
   
    
CopiarALocal:
    m_HayQueCopiarALocal = EnumSino.Sí
    fso.CopyFile m_URLEjecutableRemoto, m_URLEjecutableLocal, True
    fso.CopyFile m_URLIconoRemoto, m_URLIconoLocal, True
    fso.CopyFile m_URLbmpRemoto, m_URLbmpLocal, True
    fso.CopyFile m_URLIniRemoto, m_URLIniLocal, True
    Set wks = DBEngine.Workspaces(0)
    Set m_db = wks.OpenDatabase(m_URLEjecutableLocal, True, False, "MS Access;PWD=" & m_Pass)
    AlterarPropriedad m_URLEjecutableLocal, "AppIcon", dbText, m_URLIconoRemoto, m_db, m_Pass, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_db.NewPassword m_Pass, ""
    m_db.Close
    Set m_db = Nothing
    wks.Close
    Set wks = Nothing
EjecutarLocal:
    
 
    m_ComandoEjecutable = Chr(34) & m_ObjEntorno.URLEjecutableAccess & Chr(34) & " "
    m_ComandoEjecutable = m_ComandoEjecutable & Chr(34) & m_URLEjecutableLocal & Chr(34) & " " & " /cmd " & m_CorreoUsuarioEnLogin
    EjecutarShelllanzar m_ComandoEjecutable, EnumSino.No, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Application.Screen.ActiveForm.AllowEdits = True
    If m_ObjUsuarioConectadoLogin.MantenerLanzaderaAbierta = False Then
        
        DoCmd.Quit acQuitSaveAll
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método LanzarCondorPrueba ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Application.Screen.ActiveForm.AllowEdits = True
    On Error Resume Next
    If Not m_db Is Nothing Then
        m_db.Close
        Set m_db = Nothing
    End If
    If Not wks Is Nothing Then
        wks.Close
        Set wks = Nothing
    End If
    
End Function
Public Function setVersiones( _
                                p_Form As Form, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_ID As String
    Dim m_resto As String
    Dim m_objAplicacion As Aplicacion
    Dim m_Ver As String
    Dim ctl As Control
    Dim m_Titulo As String
    
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    
    If p_Form Is Nothing Then
        p_Error = "No hay un formulario activo"
        Err.Raise 1000
    End If
    If m_ObjEntorno Is Nothing Then
        Exit Function
    End If
    If m_ObjEntorno.ColAplicaciones Is Nothing Then
        Exit Function
    End If
    Set m_Col = m_ObjEntorno.ColAplicaciones
    For Each ctl In p_Form.Controls
        m_Titulo = ""
        If ctl.ControlType = acLabel Then
            If InStr(1, ctl.Tag, ";") <> 0 Then
                dato = Split(ctl.Tag, ";")
                m_ID = dato(0)
                m_resto = dato(1)
                If IsNumeric(m_ID) Then
                    If m_Col.Exists(CStr(m_ID)) Then
                        Set m_objAplicacion = m_Col(CStr(m_ID))
                    Else
                        Set m_objAplicacion = Nothing
                    End If
                    If Not m_objAplicacion Is Nothing Then
                        If m_resto = "S" Then
                            m_Titulo = "Ver. Servidor: " & m_objAplicacion.VersionServidor
                        ElseIf m_resto = "U" Then
                            m_Titulo = "Ver. Usuario: " & m_objAplicacion.VersionUsuario
                        Else
                            m_Titulo = "##.##.####"
                        End If
                    Else
                        m_Titulo = "##.##.####"
                    End If
                  ctl.Caption = m_Titulo
                    
                    
                End If
            End If
            
        End If
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método setVersiones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Public Function getBoton(frm As Form, p_Tag As String, Optional ByRef p_Error) As CommandButton
    
    Dim ctl As Control
    On Error GoTo errores
    For Each ctl In frm.Controls
        If ctl.ControlType = acCommandButton Then
            If Nz(ctl.Tag, "") = p_Tag Then
                Set getBoton = ctl
                Exit Function
            End If
        End If
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getBoton ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Private Function getListImagen( _
                                ByRef p_Imagelist As ImageList, _
                                p_i As Single) As ImageList

    On Error Resume Next
    Set getListImagen = p_Imagelist.ListImages(p_i)
    
End Function

Public Function AddImagen( _
                            ByRef p_Imagelist As ImageList, _
                            p_i As Single, _
                            p_Key As String, _
                            url As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim m_ObjListImagen As ListImage
    
    On Error GoTo errores
    
    
    Set m_ObjListImagen = getListImagen1(p_Imagelist, p_i)
    If m_ObjListImagen Is Nothing Then
        p_Imagelist.ListImages.Add p_i, p_Key, LoadPicture(url)
    Else
        p_Imagelist.ListImages.Remove p_i
        p_Imagelist.ListImages.Add p_i, p_Key, LoadPicture(url)
    End If
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AddImagen ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
    
End Function
Public Function getListImagen1( _
                                ByRef p_Imagelist As ImageList, _
                                p_i As Single, _
                                Optional ByRef p_Error As String _
                                ) As ListImage
    On Error Resume Next
    Set getListImagen1 = p_Imagelist.ListImages(p_i)
    
    
                                
End Function
Public Function RellenarListaImagenes( _
                                        ByRef p_Imagelist As ImageList, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_ID As Variant
    Dim m_objAplicacion As Aplicacion
    Dim m_URLIcono As String
    Dim m_URLDirIconoArbol As String
    Dim i As Single
    
    
    On Error GoTo errores
    
    
    
    i = 1
    For Each m_ID In m_ObjEntorno.ColAplicaciones
        If CStr(m_ID) = "1" Then GoTo siguiente
        Set m_objAplicacion = m_ObjEntorno.ColAplicaciones(m_ID)
        If m_objAplicacion.NombreBase = "Registro" Then
            Set m_objAplicacion = Nothing
            GoTo siguiente
        End If
        If m_objAplicacion.EnPruebasCalculado = EnumSino.Sí Then
            Set m_objAplicacion = Nothing
            GoTo siguiente
        End If
        
        m_URLIcono = m_objAplicacion.URLIconoParaArbol
        p_Error = m_objAplicacion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If fso.FileExists(m_URLIcono) Then
            'If Not FSO.FileExists(m_URLIcono) Then Stop
            AddImagen p_Imagelist, i, m_objAplicacion.NombreBase, m_URLIcono
        End If
        
        
        
        Set m_objAplicacion = Nothing
        i = i + 1
siguiente:
    Next
    m_URLDirIconoArbol = m_ObjEntorno.URLDirIconosArbol
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_URLIcono = m_URLDirIconoArbol & "folder.ico"
    If Not fso.FileExists(m_URLIcono) Then
        p_Error = "No se puede cargar el icono de carpeta abierta"
        Err.Raise 1000
    End If
    AddImagen p_Imagelist, i, fso.GetBaseName(m_URLIcono), m_URLIcono
    
    i = i + 1
    m_URLIcono = m_URLDirIconoArbol & "folder closed.ico"
    If Not fso.FileExists(m_URLIcono) Then
        p_Error = "No se puede cargar el icono de carpeta cerreada"
        Err.Raise 1000
    End If
    AddImagen p_Imagelist, i, fso.GetBaseName(m_URLIcono), m_URLIcono
    
    i = i + 1
    m_URLIcono = m_URLDirIconoArbol & "wmpIcono.ico"
    If Not fso.FileExists(m_URLIcono) Then
        p_Error = "No se puede cargar el icono del video"
        Err.Raise 1000
    End If
    AddImagen p_Imagelist, i, fso.GetBaseName(m_URLIcono), m_URLIcono
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarListaImagenes ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CargarArbol( _
                            ByRef p_Arbol As Object, _
                            ByRef p_Imagelist As Object, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_ID As Variant
    Dim m_IDVideo As Variant
    Dim m_objVideo As Video
    Dim m_objAplicacion As Aplicacion
    Dim m_NombreIcono As String
    Dim m_Key As String
    Dim m_KeyPadre As String
    Dim m_TituloNodo As String
    Dim m_Nodo As Object
    
    Dim ColVideos As Scripting.Dictionary
    Dim m_NombreIconoWMP As String
    Dim m_Form As Form
    
    
    On Error GoTo errores
    
   
    RellenarListaImagenes p_Imagelist, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    With p_Arbol
        .Style = 7 'tvwTreelinesPlusMinusText
        .MousePointer = ccArrow
        .LineStyle = 1 'tvwRootLines
        .LabelEdit = tvwManual
        .ImageList = p_Imagelist
        .BorderStyle = ccFixedSingle
        .Appearance = ccFlat
        .OLEDragMode = ccOLEDragManual
        .OLEDropMode = ccOLEDropNone
        .Indentation = 300
        .PathSeparator = "\"
        .HideSelection = False
        .Sorted = False
        .FullRowSelect = True
        .Enabled = True
        .Checkboxes = False
        .SingleSel = True
        .Scroll = True
        .HotTracking = False
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Refresh
        .Nodes.Clear
        
    End With
   
    If m_ObjEntorno.ColAplicaciones Is Nothing Then
        Exit Function
    End If
    m_NombreIconoWMP = "wmpIcono"
   
    For Each m_ID In m_ObjEntorno.ColAplicaciones
        If CStr(m_ID) = "1" Then
            GoTo siguiente
        End If
        Set m_objAplicacion = m_ObjEntorno.ColAplicaciones(CStr(m_ID))
        If m_objAplicacion.EnPruebasCalculado = EnumSino.Sí Then
            Set m_objAplicacion = Nothing
            GoTo siguiente
        End If
        Set ColVideos = m_objAplicacion.ColVideosDeAyuda
        p_Error = m_objAplicacion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
        m_Key = "A" & CStr(m_ID)
        m_TituloNodo = m_objAplicacion.NombreAplicacion
        m_NombreIcono = m_objAplicacion.NombreIconoArbol
        If m_NombreIcono = "" Then
            m_Key = ""
            m_KeyPadre = ""
            Set ColVideos = Nothing
            Set m_objAplicacion = Nothing
            GoTo siguiente
        End If
        On Error Resume Next
        Set m_Nodo = p_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error GoTo errores
        If Not ColVideos Is Nothing Then
            m_KeyPadre = m_Key
            
            m_NombreIcono = m_NombreIconoWMP
            For Each m_IDVideo In ColVideos
                Set m_objVideo = ColVideos(m_IDVideo)
                m_Key = "V" & m_objVideo.IDVideo
                m_TituloNodo = m_objVideo.Titulo
                On Error Resume Next
                Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
                If Err.Number <> 0 Then
                    Err.Clear
                End If
                On Error GoTo errores
                
siguienteVideo:
                Set m_objVideo = Nothing
            Next
        End If
        m_Key = ""
        m_KeyPadre = ""
        Set ColVideos = Nothing
        Set m_objAplicacion = Nothing
siguiente:
    Next
    p_Arbol.Refresh
    CargarArbol = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbol ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function Seleccionar( _
                            p_EsArchivo As Boolean, _
                            Optional p_Titulo As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_ObjfDialog As Object
    Dim varFile As Variant
    
    On Error GoTo errores
    
    If p_Titulo = "" Then
        p_Titulo = "Seleccione el archivo"
    End If
    If p_EsArchivo = True Then
        Set m_ObjfDialog = Application.FileDialog(msoFileDialogFilePicker)
    Else
        Set m_ObjfDialog = Application.FileDialog(msoFileDialogFolderPicker)
    End If
    With m_ObjfDialog
        .Show
        If p_EsArchivo Then
            .AllowMultiSelect = False
            .InitialFileName = m_ObjEntorno.URLUltimoArchivo
            .title = p_Titulo
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
        End If
        For Each varFile In .SelectedItems
            Seleccionar = CStr(varFile)
        Next
    End With
    If p_EsArchivo Then
        m_ObjEntorno.URLUltimoArchivo = CStr(varFile)
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Seleccionar ha producido el error : " & vbNewLine & Err.Description
    End If
    
End Function
Public Function EliminarNodo( _
                            ByRef p_Arbol As MSComctlLib.TreeView, _
                            Optional ByRef p_Nodo As MSComctlLib.Node, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim NodoPadre As MSComctlLib.Node
    On Error GoTo errores
    
    If p_Nodo Is Nothing Then
        Set p_Nodo = NodoActivo
    End If
    Set NodoPadre = p_Nodo.Parent
    If NodoPadre Is Nothing Then
        p_Error = "No se puede eliminar un nodo raíz"
        Err.Raise 1000
    End If
    p_Arbol.Nodes.Remove p_Nodo.Index
    'p_Arbol.Nodes(p_Nodo.Index).Remove
    p_Arbol.Refresh
    p_Arbol.Nodes(NodoPadre.key).Selected = True
    Form_FormVideosGestion.Arbol_NodeClick NodoPadre
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarNodo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function CrearNodo( _
                            ByRef p_Arbol As MSComctlLib.TreeView, _
                            ByRef p_NodoPadre As MSComctlLib.Node, _
                            ByRef p_ObjVideo As Video, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_Nodo As MSComctlLib.Node
    Dim m_KeyPadre As String
    Dim m_Key As String
    Dim m_TituloNodo As String
    Dim m_NombreIconoWMP As String
    
    On Error GoTo errores
    
    
    If p_NodoPadre Is Nothing Then
        Set p_NodoPadre = NodoActivo
        If p_NodoPadre Is Nothing Then
            p_Error = "No se puede crear un nodo de otro que no sea raíz"
            Err.Raise 1000
        End If
    End If
    If Not p_NodoPadre.Parent Is Nothing Then
        p_Error = "No se puede crear un nodo de otro que no sea raíz"
        Err.Raise 1000
    End If
    m_KeyPadre = p_NodoPadre.key
    m_Key = "V" & p_ObjVideo.IDVideo
    m_TituloNodo = p_ObjVideo.Titulo
    m_NombreIconoWMP = "wmpIcono"
    Set m_Nodo = p_Arbol.Nodes.Add(m_KeyPadre, tvwChild, m_Key, m_TituloNodo, m_NombreIconoWMP)
    
    p_Arbol.Refresh
    p_Arbol.Nodes(m_Nodo.key).Selected = True
    Form_FormVideosGestion.Arbol_NodeClick m_Nodo
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CrearNodo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function NombreArchivoRepetidoEnCarpeta( _
                                                p_NombreArchivo As String, _
                                                p_URLCarpeta As String, _
                                                Optional ByRef p_Error As String _
                                                ) As EnumSino
    
    
    Dim fichero As File
    
    On Error GoTo errores
    
    If p_NombreArchivo = "" Then
        p_Error = "No Se ha indicado el nombre del archivo"
        Err.Raise 1000
    End If
    If Not fso.FolderExists(p_URLCarpeta) Then
        p_Error = "No es alcanzable la carpeta"
        Err.Raise 1000
    End If
    If Right(p_URLCarpeta, 1) <> "\" Then
        p_URLCarpeta = p_URLCarpeta & "\"
    End If
    If fso.FileExists(p_URLCarpeta & p_NombreArchivo) Then
        NombreArchivoRepetidoEnCarpeta = EnumSino.Sí
    Else
        NombreArchivoRepetidoEnCarpeta = EnumSino.No
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método NombreArchivoRepetidoEnCarpeta ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function ExportarAExcel( _
                                ByRef lst As ListBox, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim m_Valor As String
    Dim wbLibro As Object
    Dim wbHoja As Object
    Dim intFila As Integer
    Dim fila As Integer
    Dim columna As Integer
    Dim m_URLExcel As String
    Dim i As Long
    Dim j As Long
    
    Dim AppExcel As Object
    Dim m_Encabezado As String
    Dim m_NombreCampo As Variant
    
    On Error GoTo errores
    If lst.ListCount = 1 Then
        p_Error = "Sin datos que exportar"
        Err.Raise 1000
    End If
    m_URLExcel = m_ObjEntorno.URLDirLocal & fso.GetBaseName(fso.GetTempName) & ".xlsx"
    If fso.FileExists(m_URLExcel) Then
        If FicheroAbierto(m_URLExcel) Then
            p_Error = "Cierre la consulta anterior"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLExcel, True
    End If
    For i = 0 To lst.ListCount - 1
        If m_Encabezado = "" Then
            m_Encabezado = lst.Column(i, 0)
        Else
            m_Encabezado = m_Encabezado & ";" & lst.Column(i, 0)
        End If
    Next
    
    dato = Split(m_Encabezado, ";")
    Set AppExcel = CreateObject("Excel.Application")
    AppExcel.Visible = False
    Set wbLibro = AppExcel.Workbooks.Add
    wbLibro.SaveAs m_URLExcel
    Set wbHoja = wbLibro.Worksheets(1)
    With wbHoja
        intFila = 1
        
        columna = 1
        For i = 0 To UBound(dato)
            .Cells(intFila, columna).value = dato(i)
            .range(.Cells(intFila, columna), .Cells(intFila, columna)).Font.Bold = True
            columna = columna + 1
        Next
        intFila = intFila + 1
        For i = 1 To lst.ListCount - 1
            columna = 1
            For j = 0 To UBound(dato)
                m_Valor = Nz(lst.Column(j, i), "")
                If IsDate(m_Valor) Then
                    .range(.Cells(intFila, columna), .Cells(intFila, columna)).NumberFormat = "dd/mm/yyyy"
                    .Cells(intFila, columna).value = CDate(m_Valor)
                Else
                    .Cells(intFila, columna).value = m_Valor
                End If
                
                columna = columna + 1
            Next
            
            intFila = intFila + 1
            
        Next
    End With
   
    
    With wbHoja.range("A1").CURRENTREGION
        .horizontalAlignment = xlGeneral
        .verticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .shrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    wbHoja.Cells.EntireColumn.AutoFit
    wbLibro.Close True
    Set wbLibro = Nothing
    AppExcel.Quit
    Set AppExcel = Nothing
    
    
    ExportarAExcel = m_URLExcel
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ExportarAExcel ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not wbLibro Is Nothing Then
        wbLibro.Close False
        Set wbLibro = Nothing
    End If
    If Not AppExcel Is Nothing Then
        AppExcel.Quit
        Set AppExcel = Nothing
    End If
    
End Function
Public Function Avance( _
                        p_Linea As Variant, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    
    Dim m_Form As Form
    
    On Error GoTo errores
    Set m_Form = Screen.ActiveForm
    On Error Resume Next
    Set lbl = m_Form.lblEstado
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo errores
    If lbl.Visible = False Then
        lbl.Visible = True
    End If
    VBA.DoEvents
    lbl.Caption = p_Linea
    VBA.DoEvents
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Avance ha devuelto el error: " & Err.Description
    End If
End Function
    
Public Function AvanceCerrar() As String
    
    
    
    
    On Error Resume Next
    Application.Screen.ActiveForm.Controls("lblEstado").Visible = False
    
    Exit Function
    
End Function


Public Function AlterarPropriedad( _
                                    p_URLEjecutableUsuario As String, _
                                    p_PropName As String, _
                                    p_PropType As Variant, _
                                    p_PropValue As Variant, _
                                    Optional p_db As DAO.Database, _
                                    Optional p_pass As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    '******************************************************************
    'Activa/Desactiva por software el acceso privilegiado con la tecla SHIFT
    'AlterarPropriedade "AllowBypassKey", dbBoolean, True --->>>>> Activa
    'AlterarPropriedade "AllowBypassKey", dbBoolean, False --->>>>> Desactiva
    'Solo se puede modificar desde el formulario de control --- Triple seguridad --- Pass de admin, usuario windows de admin, usuario BD Admin
    '*************************************************************************
    
    Dim m_prp As Property
    Dim wks As DAO.Workspace
    Const conPropNoEncontradaError As Long = 3270
    On Error GoTo errores
    
    If p_db Is Nothing Then
        
        If Not fso.FileExists(p_URLEjecutableUsuario) Then
            p_Error = "Indique la base de datos"
            Err.Raise 1000
        End If
        Set wks = DBEngine.Workspaces(0)
        If p_pass <> "" Then
            Set p_db = wks.OpenDatabase(p_URLEjecutableUsuario, True, False, "MS Access;PWD=" & p_pass)
        Else
            Set p_db = wks.OpenDatabase(p_URLEjecutableUsuario, True, False)
        End If
    End If
    On Error Resume Next
    With p_db
        .Properties(p_PropName) = p_PropValue
        If Err.Number = conPropNoEncontradaError Then
            Err.Clear
            On Error GoTo errores
            Set m_prp = .CreateProperty(p_PropName, p_PropType, p_PropValue)
            .Properties.Append m_prp
        End If
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AlterarPropriedad ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function LimpiarCache( _
                            p_ColAPlicaciones As Scripting.Dictionary, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    Dim m_objAplicacion As Aplicacion
    Dim m_ID As Variant
    Dim m_URLIniLocal As String
    
    
    On Error GoTo errores
    For Each m_ID In m_ObjEntorno.ColAplicaciones
        
       'If CStr(m_ID) = "51" Then Stop
        Set m_objAplicacion = m_ObjEntorno.ColAplicaciones(m_ID)
        If m_objAplicacion.ConIconoEnLanzaderaCalculado = EnumSino.No Then
            GoTo siguiente
        End If
        m_URLIniLocal = m_objAplicacion.URLIniUsuario
        If fso.FileExists(m_URLIniLocal) Then
            fso.DeleteFile m_URLIniLocal, True
        End If
        
siguiente:
        Set m_objAplicacion = Nothing
    Next
    
    m_URLIniLocal = Environ("APPDATA") & "\" & "Aplicaciones DYSN\GESTION RIESGOS PRUEBA\Gestion_Riesgos.ini"
    If fso.FileExists(m_URLIniLocal) Then
        fso.DeleteFile m_URLIniLocal, True
    End If
    LimpiarCache = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método LimpiarCache ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Private Function getUsuarioMaquina( _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim objNetwork As Object
    On Error GoTo errores
    Set objNetwork = CreateObject("Wscript.Network")
    With objNetwork
        getUsuarioMaquina = .UserName & "|" & .computername
    End With
   
    Set objNetwork = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getMaquina( _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim flag As String
    Dim dato As Variant
    On Error GoTo errores
    flag = getUsuarioMaquina
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        getMaquina = dato(1)
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getUsuariodeMaquina( _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim flag As String
    Dim dato As Variant
    On Error GoTo errores
    flag = getUsuarioMaquina
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        getUsuariodeMaquina = dato(0)
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariodeMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function RegistroApertura( _
                                Optional ByRef p_Error As String _
                                ) As String

    Dim m_SQL As String
    Dim m_ID As String
    Dim rcdDatos As DAO.Recordset
    Dim m_TextoEnOficina As String
    Dim m_NombreUsuario As String
   
    Dim m_NombreAplicacion As String
    Dim m_VersionAplicacion As String
    On Error GoTo errores
    
    
    If Not m_ObjUsuarioConectadoLogin Is Nothing Then
        m_NombreUsuario = m_ObjUsuarioConectadoLogin.Nombre
    Else
        m_NombreUsuario = "Desconocido"
    End If
    
    If m_EnOficina = Empty Then
        m_TextoEnOficina = "NA"
    Else
        If m_EnOficina = EnumSino.Sí Then
            m_TextoEnOficina = "Sí"
        Else
            m_TextoEnOficina = "No"
        End If
    End If
    m_VersionAplicacion = m_ObjEntorno.VersionAplicacion
    If m_VersionAplicacion = "" Then
        m_VersionAplicacion = "Desconocida"
    End If
    m_NombreAplicacion = NombreAplicacionSegunConfiguracion(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_ID = DameID("TbAplicacionesAperturas", "IDApertura", getdb(), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "SELECT * FROM TbAplicacionesAperturas;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDApertura") = m_ID
            .Fields("IDAplicacion") = IDAplicacion
            .Fields("NombreUsuario") = m_NombreUsuario
            .Fields("FechaApertura") = Date
            .Fields("HoraApertura") = getHora()
            .Fields("NombreAplicacion") = m_NombreAplicacion
            .Fields("EnOficina") = m_TextoEnOficina
            .Fields("UsuarioMaquina") = getUsuariodeMaquina()
            .Fields("NombreMaquina") = getMaquina()
            .Fields("VersionAplicacion") = m_VersionAplicacion

        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistroApertura ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function RegistroCierre( _
                                Optional ByRef p_Error As String _
                                ) As String

    Dim m_SQL As String
    Dim m_NombreUsuario As String
    On Error GoTo errores
    
    If Not m_ObjUsuarioConectadoLogin Is Nothing Then
        m_NombreUsuario = m_ObjUsuarioConectadoLogin.Nombre
    Else
        m_NombreUsuario = "Desconocido"
    End If
    
    m_SQL = "UPDATE TbAplicacionesAperturas SET FechaCierre = Date(), HoraCierre = #" & getHora() & "# " & _
            "WHERE NombreUsuario='" & m_NombreUsuario & "' " & _
            "AND HoraCierre Is Null " & _
            "AND IDAplicacion=" & IDAplicacion & ";"
    getdb().Execute m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistroCierre ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function getHora() As String
    Dim xmlHttp As Object
    Dim url As String
    Dim response As String
    Dim json As Object
    Dim hora As String
    Dim dato As Variant
    Dim flag As String
    On Error GoTo errores
    ' URL de la API para obtener la hora en Madrid, España
    url = "http://worldtimeapi.org/api/timezone/Europe/Madrid"
    
    ' Crear el objeto XMLHTTP
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Hacer la solicitud a la API
    xmlHttp.Open "GET", url, False
    xmlHttp.send
    
    ' Obtener la respuesta
    response = xmlHttp.responseText
    
    ' Analizar la respuesta JSON
    Set json = JsonConverter.ParseJson(response)
    
    ' Extraer la hora de la respuesta JSON
    hora = json("datetime")
    dato = Split(hora, "T")
    flag = dato(1)
    dato = Split(flag, ".")
    hora = dato(0)
    ' Retornar la hora
    getHora = hora
    Exit Function
errores:
    getHora = Format(Now(), "hh:mm:ss")
    
End Function
Public Function MostrarHTML( _
                            p_Mensaje As ADODB.stream, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    Dim m_ObjHTML As HTML
    On Error GoTo errores
    
    Set m_ObjHTML = New HTML
    m_ObjHTML.MostrarHTML p_Mensaje, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjHTML = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MostrarHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function HTMLENTXT( _
                            Optional p_HTML As String, _
                            Optional m_mensaje As ADODB.stream, _
                            Optional p_Mostrandolo As EnumSino = EnumSino.Sí, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim F1 As Object
    Dim m_URLHTML As String
    Dim m_URLTXT As String
    Dim m_URLCompletaArchivo As String
    Dim m_Nombre As String
        
    On Error GoTo errores
    
    If p_HTML = "" And m_mensaje Is Nothing Then
        p_Error = "No se ha indicado el HTML"
        Err.Raise 1000
    End If
       
    DameUntxtYHtml m_URLTXT, m_URLHTML, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
    If p_HTML <> "" Then
        Set F1 = fso.CreateTextFile(m_URLTXT, True)
        F1.WriteLine p_HTML
        F1.Close
        fso.GetFile(m_URLTXT).Name = fso.GetBaseName(m_URLTXT) & ".html"
    Else
        m_mensaje.SaveToFile m_URLHTML
    End If
    If p_Mostrandolo = EnumSino.Sí Then
        Ejecutar Screen.ActiveForm.hWnd, "open", m_URLHTML, "", "", 1
    End If
    
    HTMLENTXT = m_URLHTML
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTMLENTXT ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
        
    
End Function
Private Function DameUntxtYHtml( _
                                ByRef p_URLTXT As String, _
                                ByRef p_URLHTML As String, _
                                Optional ByRef p_Error As String) As String
    
    Dim m_URLHTML As String
    Dim m_URLTXT As String
    Dim i As Integer
    
    
    Dim m_NombreHTML As String
    Dim m_Nombretxt As String
    
    On Error GoTo errores
    BorraHTMLs
    For i = 1 To 50
        m_Nombretxt = "HTML" & i & ".txt"
        m_NombreHTML = "HTML" & i & ".html"
        m_URLTXT = m_ObjEntorno.URLDirLocal & m_Nombretxt
        m_URLHTML = m_ObjEntorno.URLDirLocal & m_NombreHTML
        If Not fso.FileExists(m_URLHTML) And Not fso.FileExists(m_URLHTML) Then
            p_URLTXT = m_URLTXT
            p_URLHTML = m_URLHTML
            Exit Function
        End If
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameUntxtYHtml ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Private Function BorraHTMLs( _
                            Optional ByRef p_Error As String) As String
    
    Dim fichero As File
    
    On Error GoTo errores
    
    For Each fichero In fso.GetFolder(m_ObjEntorno.URLDirLocal).Files
        If fso.GetExtensionName(fichero.Path) = "html" Or fso.GetExtensionName(fichero.Path) = "htm" Then
            If Not FicheroAbierto(fichero.Path) Then
                fso.DeleteFile fichero.Path
            End If
        End If
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método BorraHTMLs ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Public Function MostrarAperturas( _
                                    p_Tipo As EnumApertura, _
                                    Optional p_IDAplicacion As String, _
                                    Optional p_NombreUsuario As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_ObjHTML As HTML
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Apertura As AplicacionApertura
    Dim m_ColResultante As Scripting.Dictionary
    Dim m_Titulo As String
    
    On Error GoTo errores
    Set m_Col = Constructor.getAperturasPorTipo(p_Tipo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_Tipo = EnumApertura.Todas Then
        
        m_Titulo = "TODAS"
    ElseIf p_Tipo = EnumApertura.HoyTodas Then
        
        m_Titulo = "TODAS " & Date
    ElseIf p_Tipo = EnumApertura.TodasAbiertas Then
        m_Titulo = "TODAS ABIERTAS"
    ElseIf p_Tipo = EnumApertura.HoyAbiertas Then
       m_Titulo = "ABIERTAS " & Date
    ElseIf p_Tipo = EnumApertura.HoyCerradas Then
         m_Titulo = "CERRADAS " & Date
    Else
        Exit Function
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    If p_IDAplicacion = "" And p_NombreUsuario = "" Then
        Set m_ColResultante = m_Col
    Else
        For Each m_ID In m_Col
            Set m_Apertura = m_Col(m_ID)
            If p_IDAplicacion <> "" Then
                If m_Apertura.IDAplicacion <> p_IDAplicacion Then
                    GoTo siguiente
                End If
            End If
            If p_NombreUsuario <> "" Then
                If m_Apertura.NombreUsuario <> p_NombreUsuario Then
                    GoTo siguiente
                End If
            End If
            If m_ColResultante Is Nothing Then
                Set m_ColResultante = New Scripting.Dictionary
                m_ColResultante.CompareMode = TextCompare
            End If
            If Not m_ColResultante.Exists(CStr(m_Apertura.IDApertura)) Then
                m_ColResultante.Add CStr(m_Apertura.IDApertura), m_Apertura
            End If
siguiente:
            Set m_Apertura = Nothing
        Next
    End If
    If m_ColResultante Is Nothing Then
        Exit Function
    End If
    Set m_ObjHTML = New HTML
    MostrarAperturas = m_ObjHTML.MostrarAperturas(m_ColResultante, m_Titulo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjHTML = Nothing
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MostrarAperturas ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Function getLabel(p_Form As Form, p_IDAplicacion As String, Optional ByRef p_Error As String) As Label
    
    Dim ctl As Control
    
    On Error GoTo errores
    For Each ctl In p_Form.Controls
        If ctl.ControlType = acLabel Then
            If Nz(ctl.Tag, p_IDAplicacion) <> 0 Then
                Set getLabel = ctl
                Exit Function
            End If
        End If
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarEtiquetaAperturas ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function ActualizarAperturas( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    Dim ctl As Control
    Dim frm1 As Form
    Dim m_Form As Form
    Dim lbl As Label
    On Error GoTo errores
    If Not FormularioAbierto("FormMenuPrincipalAdmin") Then
        Exit Function
    End If
    Set frm1 = Forms("FormMenuPrincipalAdmin")
    Set frm1 = frm1.FormDetalle.Form
    Set m_Form = frm1.FormDetalle.Form
    If m_Form Is Nothing Then
        Exit Function
    End If
    For Each ctl In m_Form.Controls
        'If ctl.Name = "lblAperturasAgedysTec" Then Stop
        If InStr(1, ctl.Name, "lblAperturas") <> 0 Then
            If IsNumeric(Nz(ctl.Tag, "")) Then
                On Error Resume Next
                Set lbl = m_Form.Controls(ctl.Name)
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo errores
                    GoTo siguiente
                End If
                RellenarEtiquetaAperturas lbl, , , p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
        End If
siguiente:
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarAperturas ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Function RellenarEtiquetaAperturas( _
                                            Optional lbl As Label, _
                                            Optional p_Form As Form, _
                                            Optional p_IDAplicacion As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_ColAperturas As Scripting.Dictionary
    Dim m_ID As String
    
    Dim m_NumeroAperturas As Long
    On Error GoTo errores
    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado <> EnumSino.Sí Then
        Exit Function
    End If
    If Not lbl Is Nothing Then
        If Not IsNumeric(Nz(lbl.Tag, "")) Then
            Exit Function
        End If
        m_ID = lbl.Tag
    Else
        If p_Form Is Nothing Or Not IsNumeric(p_IDAplicacion) Then
            Exit Function
        End If
        Set lbl = getLabel(p_Form, p_IDAplicacion, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If lbl Is Nothing Then
            Exit Function
        End If
        m_ID = p_IDAplicacion
    End If
    'If lbl.Name = "lblAperturasAgedysTec" Then Stop
    Set m_ColAperturas = Constructor.getUsuariosAplicacionAbierta(m_ID, p_Error)
    
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ColAperturas Is Nothing Then
        m_NumeroAperturas = m_ColAperturas.Count
    End If
    lbl.Caption = m_NumeroAperturas
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarEtiquetaAperturas ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function CerrarAperturas( _
                                    Optional p_SoloDiasAnteriores As EnumSino, _
                                    Optional ByRef p_Error As String _
                                    ) As String

    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    
    
    On Error GoTo errores
    
    If p_SoloDiasAnteriores = Empty Then p_SoloDiasAnteriores = EnumSino.Sí
    If p_SoloDiasAnteriores = EnumSino.Sí Then
        m_SQL = "SELECT * " & _
                "FROM TbAplicacionesAperturas " & _
                "WHERE FechaCierre Is Null " & _
                "AND FechaApertura<#" & Format(Date, "mm/dd/yyyy") & "#;"
                
    Else
        m_SQL = "SELECT * " & _
                "FROM TbAplicacionesAperturas " & _
                "WHERE FechaCierre Is Null;"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                .Edit
                    .Fields("FechaCierre") = Date
                    .Fields("HoraCierre") = getHora()
                    .Fields("Observaciones") = "Se cierra por Administrador desde Lanzadera el " & Now()
                .Update
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CerrarAperturas ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Function getDirectorioOneDrive(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim carpetaRaiz As Object
    Dim subCarpeta As Object
    Dim rutaEncontrada As String
    Dim encontrado As Boolean
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la carpeta raíz de C:\
    Set carpetaRaiz = fso.GetFolder("C:\")
    
    ' Inicializar variables
    encontrado = False
    rutaEncontrada = ""
    
    ' Recorrer las subcarpetas en la raíz de C:\
    For Each subCarpeta In carpetaRaiz.SubFolders
        If InStr(1, subCarpeta.Name, "OneDrive", vbTextCompare) > 0 Then
            rutaEncontrada = subCarpeta.Path
            encontrado = True
            Exit For
        
        End If
    Next subCarpeta
    
    ' Mostrar el resultado
    If encontrado Then
        getDirectorioOneDrive = rutaEncontrada
    
    End If
    
    ' Liberar objetos
    Set subCarpeta = Nothing
    Set carpetaRaiz = Nothing
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDrive ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function

Private Function getDirectorioOneDriveTelefonicaApps(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim carpeta As String
    Dim m_RutaOneDrive As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Exit Function
    End If
    carpeta = m_RutaOneDrive & "\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\"
    If Not fso.FolderExists(carpeta) Then
        Exit Function
    End If
    getDirectorioOneDriveTelefonicaApps = carpeta
    
    ' Liberar objetos
   
    
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDriveTelefonicaApps ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function
Private Function getDirectorioOneDriveApps(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim carpeta As String
    Dim m_RutaOneDrive As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Exit Function
    End If
    'C:\OneDrive\OneDrive - Telefonica\00LABORAL\Aplicaciones PpD
    carpeta = m_RutaOneDrive & "\OneDrive - Telefonica\00LABORAL\Aplicaciones PpD\"
    If Not fso.FolderExists(carpeta) Then
        Exit Function
    End If
    getDirectorioOneDriveApps = carpeta
    
    ' Liberar objetos
   
    
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDriveApps ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function getRutaAplicacionesLocal(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim m_RutaOneDrive As String
    Dim m_RutaOneDriveTelefonica As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    m_RutaOneDriveTelefonica = getDirectorioOneDriveTelefonicaApps(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If fso.FolderExists(m_RutaOneDriveTelefonica) Then
        Set fso = Nothing
        getRutaAplicacionesLocal = m_RutaOneDriveTelefonica
        Exit Function
    End If
    
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Set fso = Nothing
        Exit Function
    End If
   
    getRutaAplicacionesLocal = m_RutaOneDrive
    
    ' Liberar objetos
   
    
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRutaAplicacionesLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function



Public Function GetConfig(sClave As String, Optional vDefault As Variant = Null) As Variant
    Dim rs As DAO.Recordset
    Set rs = getdb().OpenRecordset( _
        "SELECT Valor, TipoDato FROM TbConfiguracion WHERE Clave='" & sClave & "'", _
        dbOpenSnapshot)
    
    If rs.EOF Then
        GetConfig = vDefault
    Else
        Select Case rs!TipoDato
            Case "Boolean"
                GetConfig = (rs!Valor = "1" Or LCase(rs!Valor) = "true")
            Case "Integer"
                GetConfig = CInt(rs!Valor)
            Case "Long"
                GetConfig = CLng(rs!Valor)
            Case Else
                GetConfig = rs!Valor
        End Select
    End If
    
    rs.Close
    Set rs = Nothing
End Function



