Attribute VB_Name = "FUNCIONES UTILES"

Option Compare Database
Option Explicit
Private s_contadorPasos As Long
' ==========================================================================
' MÓDULO: FUNCIONES UTILES.bas (CENTRALIZADO CON TODAS LAS API DE WINDOWS)
' ==========================================================================

' --- Declaraciones API de Windows (ÚNICA FUENTE DE VERDAD) ---
#If Win64 = 1 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare PtrSafe Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As LongPtr)
    'Public Declare PtrSafe Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
    Public Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
    Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
    Public Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    Public Declare PtrSafe Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
    Public Declare PtrSafe Function IsIconic Lib "user32.dll" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Public Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
    Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Boolean
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
   ' Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
    Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
    Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Public Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
    Public Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
    Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function EmptyClipboard Lib "user32" () As Long
    Public Declare Function CloseClipboard Lib "user32" () As Long
    Public Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As Long) As Long
    Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Boolean
#End If
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STATUS_PENDING = &H103&

Public Function getWorkspace() As DAO.Workspace
    On Error Resume Next
    If g_wsCondor Is Nothing Then
        Set g_wsCondor = DBEngine.CreateWorkspace("CondorWS", "admin", "")
    End If
    Set getWorkspace = g_wsCondor
End Function

Public Function getdbLanzadera() As DAO.Database
    On Error GoTo Errores
    
    ' Comprobación de estado: Si la conexión no existe o está cerrada, se restablece.
    Call CheckAndReconnect(g_dbLanzadera, "Lanzadera")
    
    Set getdbLanzadera = g_dbLanzadera
    Exit Function
Errores:
    If Err.Number = 3045 Then ' Archivo ya en uso
        MsgBox "No se puede iniciar CONDOR." & vbNewLine & vbNewLine & _
               "La base de datos de usuarios (Lanzadera_Datos.accdb) ya está abierta en modo exclusivo.", vbCritical, "Error de Acceso a Datos"
        Application.Quit
    Else ' Otro tipo de error
        Dim errObj As New CondorError
        errObj.Create Err.Number, "No se pudo conectar a la base de datos de Lanzadera. " & Err.description, "FUNCIONES UTILES.getdbLanzadera"
        errObj.Raise
    End If
End Function

Public Function getdb() As DAO.Database
    On Error GoTo Errores

    ' Test mode is an explicit safety override: when enabled, every normal
    ' getdb() caller is routed to the validated sandbox/backend configured by
    ' TestHelper.ForceLocalBackend. Production behavior below is unchanged.
    If m_TestingMode Then
        Dim testPassword As String

        If Len(Trim$(m_BackendSandboxURL)) = 0 Then
            Err.Raise 513, "FUNCIONES UTILES.getdb", "TESTS BLOCKED: m_TestingMode=True but m_BackendSandboxURL is empty"
        End If

        ' PR2 (Spec-005/008): prefer the test-side cache m_BackendSandboxPassword
        ' so that test routing does not depend on m_PasswordBackend, which
        ' ResetTestSession must NOT clear (production cache).
        testPassword = m_BackendSandboxPassword
        If Len(testPassword) = 0 Then testPassword = m_PasswordBackend
        If Len(testPassword) = 0 Then testPassword = GetPasswordDB()

        ' Cache safety (Spec-008): if g_dbCondor points anywhere other than
        ' the asserted sandbox (production, previous-session test, or stale
        ' production path), close it and reopen. Never return a cached
        ' connection from a previous session.
        If g_dbCondor Is Nothing Then
            Set g_dbCondor = getWorkspace().OpenDatabase(m_BackendSandboxURL, False, False, ";PWD=" & testPassword)
        ElseIf StrComp(g_dbCondor.name, m_BackendSandboxURL, vbTextCompare) <> 0 Then
            On Error Resume Next
            g_dbCondor.Close
            Set g_dbCondor = Nothing
            On Error GoTo Errores
            Set g_dbCondor = getWorkspace().OpenDatabase(m_BackendSandboxURL, False, False, ";PWD=" & testPassword)
        End If

        Set getdb = g_dbCondor
        Exit Function
    End If

    ' Comprobación de estado: Si la conexión no existe o está cerrada, se restablece.
    Call CheckAndReconnect(g_dbCondor, "Condor")

    Set getdb = g_dbCondor
    Exit Function
Errores:
    If Err.Number = 3045 Then ' Archivo ya en uso
        MsgBox "No se puede iniciar CONDOR." & vbNewLine & vbNewLine & _
               "La base de datos principal (Condor_Datos.accdb) ya está abierta en modo exclusivo.", vbCritical, "Error de Acceso a Datos"
        Application.Quit
    Else ' Otro tipo de error
        Dim errObj As New CondorError
        errObj.Create Err.Number, "No se pudo conectar a la base de datos principal de CONDOR. " & Err.description, "FUNCIONES UTILES.getdb"
        errObj.Raise
    End If
End Function

' -----------------------------------------------------------------------------
' PasswordDB: Abstraction layer para el password de la BD
'   Lee desde Entorno.PasswordDB (INI) con fallback hardcodeado legacy
' -----------------------------------------------------------------------------
Public Function GetPasswordDB() As String
    On Error GoTo Errores
    Dim objEntorno As New entorno
    GetPasswordDB = objEntorno.PasswordDB
    Exit Function
Errores:
    ' Fallback hardcoded legacy
    GetPasswordDB = "dpddpd"
End Function

Public Function getdbCorreo() As DAO.Database
    On Error GoTo Errores
    
    ' Comprobación de estado: Si la conexión no existe o está cerrada, se restablece.
    Call CheckAndReconnect(g_dbCorreos, "Correos")
    
    Set getdbCorreo = g_dbCorreos
    Exit Function
Errores:
    If Err.Number = 3045 Then ' Archivo ya en uso
        MsgBox "No se puede completar la operación." & vbNewLine & vbNewLine & _
               "La base de datos de correos (Correos_datos.accdb) ya está abierta en modo exclusivo.", vbCritical, "Error de Acceso a Datos"
        Application.Quit
    Else ' Otro tipo de error
        Dim errObj As New CondorError
        errObj.Create Err.Number, "No se pudo conectar a la base de datos de Correos. " & Err.description, "FUNCIONES UTILES.getdbCorreo"
        errObj.Raise
    End If
End Function

Public Function getdbExpedientes() As DAO.Database
    On Error GoTo Errores
    
    ' Comprobación de estado: Si la conexión no existe o está cerrada, se restablece.
    Call CheckAndReconnect(g_dbExpedientes, "Expedientes")
    
    Set getdbExpedientes = g_dbExpedientes
    Exit Function
Errores:
    If Err.Number = 3045 Then ' Archivo ya en uso
        MsgBox "No se puede iniciar CONDOR." & vbNewLine & vbNewLine & _
               "La base de datos de expedientes (Expedientes_datos.accdb) ya está abierta en modo exclusivo.", vbCritical, "Error de Acceso a Datos"
        Application.Quit
    Else ' Otro tipo de error
        Dim errObj As New CondorError
        errObj.Create Err.Number, "No se pudo conectar a la base de datos de Expedientes. " & Err.description, "FUNCIONES UTILES.getdbExpedientes"
        errObj.Raise
    End If
End Function
Public Function getdbNoConformidades() As DAO.Database
    On Error GoTo Errores
    
    ' Comprobación de estado: Si la conexión no existe o está cerrada, se restablece.
    Call CheckAndReconnect(g_dbNoConformidades, "NoConformidades")
    
    Set getdbNoConformidades = g_dbNoConformidades
    Exit Function
Errores:
    If Err.Number = 3045 Then ' Archivo ya en uso
        MsgBox "No se puede iniciar CONDOR." & vbNewLine & vbNewLine & _
               "La base de datos de expedientes (NoConformidades_datos.accdb) ya está abierta en modo exclusivo.", vbCritical, "Error de Acceso a Datos"
        Application.Quit
    Else ' Otro tipo de error
        Dim errObj As New CondorError
        errObj.Create Err.Number, "No se pudo conectar a la base de datos de No Conformidades. " & Err.description, "FUNCIONES UTILES.getdbNoConformidades"
        errObj.Raise
    End If
End Function

Private Sub CheckAndReconnect(ByRef dbObject As DAO.Database, ByVal dbType As String)
    Dim needsConnection As Boolean
    Dim dummy As String
    Dim dbURL As String
    Dim dbPassword As String
    Dim dbName As String

    On Error GoTo Errores

    needsConnection = False

    If dbObject Is Nothing Then
        needsConnection = True
    Else
        On Error Resume Next
        dummy = dbObject.name
        If Err.Number <> 0 Then
            needsConnection = True
            Err.Clear
        End If
        On Error GoTo Errores
    End If

    ' --- PR2 (Spec-005/008): top-level test-mode routing ----------------------
    ' When m_TestingMode is active, EVERY getdb*() family member (Lanzadera,
    ' Condor, Correos, Expedientes, NoConformidades) must open the asserted
    ' sandbox backend. This closes the RED atom on
    ' Test_HarnessV24_BeginTestSession_Contract where EVE -> getdbLanzadera
    ' was resolving to the production LANZADERA path. The transitional shim
    ' here is for the lifecycle probe only; project AGENTS.md forbids
    ' business code from calling getdb*() outside getdb().
    If needsConnection And m_TestingMode Then
        If Len(Trim$(m_BackendSandboxURL)) = 0 Then
            Err.Raise 513, "CheckAndReconnect", _
                "TESTS BLOCKED: m_TestingMode=True but m_BackendSandboxURL is empty"
        End If
        dbURL = m_BackendSandboxURL
        dbPassword = m_BackendSandboxPassword
        If Len(dbPassword) = 0 Then dbPassword = m_PasswordBackend
        GoTo OpenResolvedDb
    End If

    If needsConnection Then
        Select Case dbType
            Case "Lanzadera"
                 If Application.TempVars("DatosEnLocal") = "Sí" Then dbURL = m_URLRutaAplicacionesLocal & "000datoslocal\Lanzadera_Datos.accdb" Else dbURL = m_URLRutaAplicacionesRemotas & "0Lanzadera\Lanzadera_Datos.accdb"
            Case "Condor"
            Select Case Application.TempVars("BackendActivo")
                Case "TEST"
                    dbURL = Application.TempVars("BackendTest")
                Case "SANDBOX"
                    dbURL = Application.TempVars("BackendSandbox")
                Case "PROD"
                    dbURL = Application.TempVars("BackendProduccion")
                Case Else
                    ' Fallback legacy behavior for unspecified backend
                    dbName = IIf(Application.TempVars("EnPruebas") = "Sí", "Condor_Datos_PRUEBAS.accdb", "Condor_Datos.accdb")
                    If Application.TempVars("DatosEnLocal") = "Sí" Then dbURL = m_URLRutaAplicacionLocal & dbName Else dbURL = m_URLRutaAplicacionRemota & dbName
            End Select
            Case "Correos"
                If Application.TempVars("DatosEnLocal") = "Sí" Then dbURL = m_URLRutaAplicacionesLocal & "000datoslocal\Correos_datos.accdb" Else dbURL = m_URLRutaAplicacionesRemotas & "00Recursos\Correos_datos.accdb"
            Case "Expedientes"
                 If Application.TempVars("DatosEnLocal") = "Sí" Then dbURL = m_URLRutaAplicacionesLocal & "000datoslocal\Expedientes_datos.accdb" Else dbURL = m_URLRutaAplicacionesRemotas & "EXPEDIENTES\Expedientes_datos.accdb"
            Case "NoConformidades"
                If Application.TempVars("DatosEnLocal") = "Sí" Then dbURL = m_URLRutaAplicacionesLocal & "000datoslocal\NoConformidades_Datos.accdb" Else dbURL = m_URLRutaAplicacionesRemotas & "No Conformidades\NoConformidades_Datos.accdb"
            Case Else
                Err.Raise 513, "CheckAndReconnect", "Tipo de base de datos desconocido: '" & dbType & "'"
        End Select

        ' --- CORRECCIÓN CRÍTICA DE BLOQUEO ---
        ' 1. Usamos getWorkspace() en lugar de DBEngine.Workspaces(0)
        ' 2. Corregimos la cadena de conexión a ";PWD=..." (Formato estándar DAO)
OpenResolvedDb:
        If Len(dbPassword) = 0 Then dbPassword = GetPasswordDB()
        Set dbObject = getWorkspace().OpenDatabase(dbURL, False, False, ";PWD=" & dbPassword)
    End If

LimpiarYSalir:
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "FUNCIONES UTILES.CheckAndReconnect"
    errObj.AddToCallStack "Intentando (re)conectar a la base de datos: " & dbType
    errObj.Raise
End Sub
Public Sub CerrarConexionesGlobales()
    On Error Resume Next
    If Not g_dbCondor Is Nothing Then g_dbCondor.Close: Set g_dbCondor = Nothing
    If Not g_dbLanzadera Is Nothing Then g_dbLanzadera.Close: Set g_dbLanzadera = Nothing
    If Not g_dbCorreos Is Nothing Then g_dbCorreos.Close: Set g_dbCorreos = Nothing
    If Not g_dbExpedientes Is Nothing Then g_dbExpedientes.Close: Set g_dbExpedientes = Nothing
    If Not g_dbNoConformidades Is Nothing Then g_dbNoConformidades.Close: Set g_dbNoConformidades = Nothing
    
    ' Cerramos el workspace al final
    If Not g_wsCondor Is Nothing Then g_wsCondor.Close: Set g_wsCondor = Nothing
End Sub


Public Sub QuitarFocoControl(ByRef frm As Form, ByVal nombreControl As String)
    On Error Resume Next
    If frm Is Nothing Then Exit Sub
    If frm.ActiveControl Is Nothing Then Exit Sub
    If frm.ActiveControl.name = nombreControl Then
        frm.SetFocus
    End If
    On Error GoTo 0
End Sub

Public Function FormularioAbierto(ByVal nombre As String) As Boolean
    Dim estado As Long
    Dim frm As Form
    
    On Error GoTo Errores
    
    On Error Resume Next
    estado = SysCmd(acSysCmdGetObjectState, acForm, nombre)
    If Err.Number <> 0 Or estado = 0 Then
        Err.Clear
        FormularioAbierto = False
        GoTo LimpiarYSalir
    End If
    On Error GoTo Errores
    
    On Error Resume Next
    Set frm = Forms(nombre)
    If Err.Number <> 0 Then
        Err.Clear
        FormularioAbierto = False
        GoTo LimpiarYSalir
    End If
    
    If Not frm.Visible Then
        FormularioAbierto = False
    Else
        FormularioAbierto = True
    End If

LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "FUNCIONES UTILES.FormularioAbierto"
    errObj.AddToCallStack "Comprobando formulario: " & nombre
    errObj.Raise
End Function

Public Sub Ajustar(ByRef frmFormulario As Form)
    Dim i As Integer
    On Error Resume Next
    If Not frmFormulario.RecordSelectors Then
        frmFormulario.InsideWidth = frmFormulario.Width
    Else
        frmFormulario.InsideWidth = frmFormulario.Width + 250
    End If
    If frmFormulario.DefaultView = 0 Then
        frmFormulario.InsideHeight = 0
        For i = 0 To 100
            frmFormulario.InsideHeight = frmFormulario.InsideHeight + frmFormulario.Section(i).Height
        Next
    End If
End Sub

' --- INICIO DE LA CORRECCIÓN: Lógica de la barra de progreso ---

' Variable estática a nivel de módulo para el contador


Public Sub ResetAvanceCounter()
    ' Responsabilidad: Pone a cero el contador de pasos del splash.
    ' Debe ser llamado desde el evento Form_Open del frmSplash.
    s_contadorPasos = 0
End Sub

Public Sub Avance(ByRef p_Linea As Variant)
    Dim frm As Form
    On Error Resume Next
    
    Set frm = Screen.ActiveForm
    If Not frm Is Nothing Then
        If Not frm.Controls("lblEstado") Is Nothing Then
            frm.Controls("lblEstado").Caption = p_Linea
        End If
    End If
    
    If FormularioAbierto("frmSplash") Then
        With Forms("frmSplash")
            Dim totalPasos As Long
            Dim anchoMaximo As Long, nuevoAncho As Long
            
            Dim tempEntorno As New entorno
            totalPasos = tempEntorno.ColItems.count
            Set tempEntorno = Nothing
            
            s_contadorPasos = s_contadorPasos + 1
            
            anchoMaximo = .lblProgresoFondo.Width
            
            If totalPasos > 0 Then
                ' Cálculo proporcional
                nuevoAncho = (s_contadorPasos / totalPasos) * anchoMaximo
                
                ' SALVAGUARDA: Asegurarse de que el nuevo ancho no supere el máximo.
                If nuevoAncho > anchoMaximo Then
                    nuevoAncho = anchoMaximo
                End If
            End If
            
            .lblProgresoBarra.Width = nuevoAncho
        End With
    End If
    
    VBA.DoEvents
End Sub
' --- FIN DE LA CORRECCIÓN ---
Public Sub AvanceCerrar()
    Dim frm As Form
    On Error Resume Next
    Set frm = Screen.ActiveForm
    If Not frm Is Nothing Then
        If Not frm.Controls("lblEstado") Is Nothing Then
            frm.Controls("lblEstado").Visible = False
        End If
    End If
End Sub

Public Sub AbrirEnLocal(ByVal p_URLFinal As String)
    Dim m_URLLocal As String
    Dim m_Hwn As LongPtr
    On Error GoTo Errores
    
    On Error Resume Next
    m_Hwn = Application.Screen.ActiveForm.hwnd
    If Err.Number <> 0 Then
        Err.Clear
        m_Hwn = 0
    End If
    On Error GoTo Errores

    If Not fso.FileExists(p_URLFinal) Then
        Err.Raise 513, , "No es accesible la ruta del archivo: " & p_URLFinal
    End If
    
    If Left(p_URLFinal, 2) = "\\" Then
        m_URLLocal = m_ObjEntorno.URLDirectorioLocal & fso.GetFileName(p_URLFinal)
        If fso.FileExists(m_URLLocal) Then
            If FicheroAbierto(m_URLLocal) Then
                Err.Raise 513, , "El archivo temporal ya está abierto. Cierre el documento y vuelva a intentarlo."
            End If
            fso.DeleteFile m_URLLocal, True
        End If
        fso.CopyFile p_URLFinal, m_URLLocal, True
    Else
        m_URLLocal = p_URLFinal
    End If
    
    Ejecutar m_Hwn, "open", m_URLLocal, "", "", 1
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "FUNCIONES UTILES.AbrirEnLocal"
    errObj.Raise
End Sub

Public Sub AbrirAyuda()
    Dim m_URLNombreArchivo As String
    On Error GoTo Errores
    
    On Error Resume Next
    If Application.Screen.ActiveForm Is Nothing Then Exit Sub
    On Error GoTo Errores
    
    m_URLNombreArchivo = m_ObjEntorno.URLDirectorioDocumentacionAyuda & "Gestor de Expedientes - Presentación y usos v1.pdf"
    If Not fso.FileExists(m_URLNombreArchivo) Then
        Err.Raise 513, , "No se ha encontrado el fichero de ayuda."
    End If
    
    Call AbrirEnLocal(m_URLNombreArchivo)
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "FUNCIONES UTILES.AbrirAyuda"
    errObj.Raise
End Sub

Public Function DameID(p_NOmbreTabla As String, p_NombreCampoID As String, ByRef p_db As DAO.Database) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim lngOrdinalMaximo As Long
    Dim dummy As String
    On Error GoTo Errores
    
    ' VALIDACIÓN CRÍTICA: Verificar que p_db sea un objeto DAO.Database válido
    ' Error 3420 "objeto no válido" puede ocurrir si VBA garbage collector liberó el objeto
    ' o si se pasó un objeto inválido (ej: Empty desde Optional ByRef mal manejado)
    If p_db Is Nothing Then
        Err.Raise 513, "DameID", "Parámetro db es Nothing. La conexión a la base de datos no está disponible."
    End If
    
    On Error Resume Next
    dummy = p_db.name
    If Err.Number <> 0 Then
        Err.Raise 513, "DameID", "El objeto Database es inválido (Error " & Err.Number & " verificando .Name). La conexión puede haber sido cerrada o liberada por el garbage collector."
    End If
    On Error GoTo Errores
    
    m_SQL = "SELECT Max(" & p_NOmbreTabla & ".[" & p_NombreCampoID & "]) AS Maximo FROM [" & p_NOmbreTabla & "];"
    Set rcdDatos = p_db.OpenRecordset(m_SQL, dbOpenSnapshot)
    
    If Not rcdDatos.EOF Then
        lngOrdinalMaximo = Nz(rcdDatos.Fields("Maximo").value, 0)
    End If
    
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameID = CStr(lngOrdinalMaximo + 1)
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "FUNCIONES UTILES.DameID"
    errObj.Raise
End Function

Function FicheroAbierto(strURLArchivo As String) As Boolean
    Dim intNumeroFichero As Integer, intNumeroError As Integer
    On Error Resume Next
    intNumeroFichero = FreeFile()
    Open strURLArchivo For Input Lock Read As #intNumeroFichero
    Close intNumeroFichero
    intNumeroError = Err.Number
    On Error GoTo 0
    Select Case intNumeroError
        Case 0: FicheroAbierto = False
        Case Else: FicheroAbierto = True
    End Select
End Function

Public Sub EliminarContenidoCarpeta(ByVal rutaCarpeta As String)
    On Error GoTo Errores
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(rutaCarpeta) Then
        fso.DeleteFile rutaCarpeta & "\*.*", True
        fso.DeleteFolder rutaCarpeta & "\*", True
    End If
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "FUNCIONES UTILES.EliminarContenidoCarpeta"
    errObj.Raise
End Sub

' ==========================================================================
' MÓDULO: FUNCIONES UTILES.bas
' ACCIÓN: Corrección de manejo de errores (Integración con CondorError)
' ==========================================================================
Public Sub LlenarComboDesdeDiccionario(ByRef cbo As ComboBox, ByVal col As Scripting.Dictionary, ByVal campoID As String, ByVal campoTexto As String, Optional ByVal conOpcionTodos As Boolean = True)
    Dim unObjeto As Object
    Dim key As Variant
    Dim rowSourceString As String
    
    On Error GoTo Errores
    
    ' Validación defensiva básica
    If cbo Is Nothing Then Err.Raise 513, , "El control ComboBox es nulo."
    
    cbo.rowSource = ""
    cbo.RowSourceType = "Value List"
    cbo.ColumnCount = 2
    cbo.BoundColumn = 1
    cbo.ColumnWidths = "0cm;8cm"
    
    If conOpcionTodos Then
        rowSourceString = "0;'(Todos)'"
    End If
    
    If Not col Is Nothing Then
        For Each key In col.Keys
            Set unObjeto = col(key)
            If rowSourceString <> "" Then rowSourceString = rowSourceString & ";"
            
            ' Usamos getters seguros o acceso directo si son clases estándar
            rowSourceString = rowSourceString & unObjeto.getPropiedad(campoID) & ";" & unObjeto.getPropiedad(campoTexto)
        Next key
    End If
    
    cbo.rowSource = rowSourceString
    If conOpcionTodos Then cbo.value = 0
    
    Exit Sub

Errores:
    ' CORRECCIÓN: Usamos CondorError para propagar en lugar de MsgBox directo
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "FUNCIONES UTILES.LlenarComboDesdeDiccionario"
    
    ' Añadimos contexto útil para saber qué combo falló
    If Not cbo Is Nothing Then
        errObj.AddToCallStack "Control: " & cbo.name
    Else
        errObj.AddToCallStack "Control: Nothing"
    End If
    
    errObj.Raise
End Sub

Public Sub ManejadorErroresFormulario(ByVal procedimiento As String, ByVal frm As Access.Form)
    ' --- 1. CAPTURA INMEDIATA DEL ERROR ORIGINAL (CRÍTICO) ---
    ' Debemos guardar estos valores ANTES de cualquier instrucción 'On Error Resume Next'
    ' o de cualquier otra operación, ya que podrían resetear el objeto Err.
    Dim inErrNum As Long
    Dim inErrDesc As String
    Dim inErrSource As String
    
    inErrNum = Err.Number
    inErrDesc = Err.description
    inErrSource = Err.source
    ' ---------------------------------------------------------

    Dim titulo As String, estilo As VbMsgBoxStyle
    Dim msg As String
    Dim ce As CondorError
    
    ' Ahora ya podemos proteger el propio manejador
    On Error Resume Next

    ' Desbloqueo de UI (Esto era lo que borraba el error antes)
    If Not frm Is Nothing Then
        If TieneControl(frm, "lblBloqueo") Then frm.Controls("lblBloqueo").Visible = False
    End If
    DoCmd.Hourglass False

    ' --- LÓGICA DE SELECCIÓN DE ERROR ---
    
    ' A. ¿Es un CondorError (error de negocio/servicio)?
    If Not g_objLastError Is Nothing Then
        Set ce = g_objLastError
        
        ' Usamos los datos del objeto CondorError
        Select Case ce.Number
            Case 513 ' VALIDACIÓN
                titulo = "Validación: Datos Requeridos"
                estilo = vbExclamation + vbOKOnly
                msg = FormatearDescripcionValidacion(ce.description)
                
            Case 514 ' ERROR OPERATIVO
                titulo = "No se puede completar la operación"
                estilo = vbCritical + vbOKOnly
                msg = ce.description
                
            Case Else ' ERROR TÉCNICO
                titulo = "Error en " & procedimiento
                estilo = vbCritical + vbOKOnly
                msg = ce.FullDescription
        End Select
        
    ' B. ¿Es un Error Estándar de VBA (ocurrido en el formulario)?
    Else
        ' Usamos las variables que capturamos al principio (inErrNum)
        ' en lugar del objeto Err (que ahora vale 0).
        
        If inErrNum = 0 Then
             ' Si llegamos aquí con 0, es que realmente no había error o se perdió antes de llamar.
             inErrDesc = "Se ha producido un error no especificado o se ha perdido el contexto."
             inErrSource = "Desconocido"
        End If
        
        titulo = "Error inesperado en " & procedimiento
        estilo = vbCritical + vbOKOnly
        msg = "Se ha producido un error no gestionado:" & vbCrLf & vbCrLf & _
              "Nº: " & inErrNum & vbCrLf & _
              "Origen: " & inErrSource & vbCrLf & _
              "Descripción: " & inErrDesc
    End If

    ' Mostrar Mensaje
    MsgBox msg, estilo, titulo

    ' Log
    Dim fullErrorDesc As String
    If Not ce Is Nothing Then
        fullErrorDesc = ce.FullDescription
    Else
        fullErrorDesc = inErrDesc
    End If
    
    ' Usamos el número capturado si no hay objeto ce
    Dim logNum As Long
    If Not ce Is Nothing Then logNum = ce.Number Else logNum = inErrNum
    
    LogErrorUI logNum, fullErrorDesc, Nz(inErrSource, ""), procedimiento

    ' LIMPIEZA FINAL
    Set g_objLastError = Nothing
    Err.Clear
End Sub
' Helper functions for error handling

Private Function GetBaseErrorNumber(ByVal errNum As Long) As Long
    ' Desenvuelve un CondorError para obtener su número de error original.
    If errNum > (vbObjectError + 20000) Then
        GetBaseErrorNumber = errNum - (vbObjectError + 20000)
    Else
        GetBaseErrorNumber = errNum
    End If
End Function

Private Function ExtraerDescripcionLimpia(ByVal fullDesc As String) As String
    ' Extrae solo la primera línea de la descripción de un CondorError.
    Dim pos As Long
    pos = InStr(1, fullDesc, "Pila de Llamadas:")
    If pos > 0 Then
        ExtraerDescripcionLimpia = Trim(Left(fullDesc, pos - 1))
        ' Limpiamos el prefijo "Error #XXX:"
        pos = InStr(1, ExtraerDescripcionLimpia, ":")
        If pos > 0 Then
            ExtraerDescripcionLimpia = Trim(Mid(ExtraerDescripcionLimpia, pos + 1))
        End If
    Else
        ExtraerDescripcionLimpia = fullDesc
    End If
End Function

' ==========================================
'  FORMATEO DE DESCRIPCIÓN DE VALIDACIÓN 513
' ==========================================
' Espera líneas del estilo:
'   • (Generales) Campo A
'   • (Propuesta) Campo B
'   • (Impacto) Campo C
'
' Si una línea no tiene prefijo "(Grupo)", se agrupa bajo "Otros".
Private Function FormatearDescripcionValidacion(ByVal descripcion As String) As String
    Dim lineas() As String, i As Long
    Dim grupos As Object ' Scripting.Dictionary late-bound
    Dim nombreGrupo As String, item As String, posCierre As Long
    Dim clave As Variant, sb As String

    On Error GoTo FinSeguro

    Set grupos = CreateObject("Scripting.Dictionary")
    grupos.CompareMode = 1 ' TextCompare

    lineas = Split(descripcion, vbCrLf)
    For i = LBound(lineas) To UBound(lineas)
        Dim ln As String
        ln = Trim(lineas(i))
        If ln <> "" Then
            nombreGrupo = "Otros"
            item = ln

            If Left$(ln, 1) = "(" Then
                posCierre = InStr(2, ln, ")")
                If posCierre > 2 Then
                    nombreGrupo = Trim(Mid$(ln, 2, posCierre - 2))
                    item = Trim(Mid$(ln, posCierre + 1))
                    ' Quitar separadores tipo "-" al inicio del item si los hubiera
                    If Left$(item, 1) = "-" Then item = Trim(Mid$(item, 2))
                End If
            ElseIf Left$(ln, 1) = "•" Then
                ' Línea ya con viñeta ? quitarla para no duplicar el punto
                item = Trim(Mid$(ln, 2))
            End If

            If Not grupos.Exists(nombreGrupo) Then
                grupos.Add nombreGrupo, New Collection
            End If
            grupos(nombreGrupo).Add item
        End If
    Next i

    ' Construir mensaje agrupado
    sb = "Faltan campos obligatorios:" & vbCrLf & vbCrLf
    For Each clave In grupos.Keys
        sb = sb & "[" & CStr(clave) & "]" & vbCrLf
        For i = 1 To grupos(clave).count
            sb = sb & "  • " & CStr(grupos(clave).item(i)) & vbCrLf
        Next i
        sb = sb & vbCrLf
    Next clave

    FormatearDescripcionValidacion = sb
    Exit Function

FinSeguro:
    ' Si algo falla en el formateo, devolvemos la descripción cruda
    FormatearDescripcionValidacion = descripcion
End Function


' ======================
'  HELPERS DE SOPORTE
' ======================
Public Function TieneControl(ByVal frm As Access.Form, ByVal nombre As String) As Boolean
    On Error GoTo Fin
    Dim ctl As Control
    Set ctl = frm.Controls(nombre)
    TieneControl = True
Fin:
End Function

Private Sub LogErrorUI(ByVal numero As Long, ByVal descripcion As String, _
                       ByVal origen As String, ByVal procedimiento As String)
    On Error GoTo Fin
    Dim db As DAO.Database
    Dim sql As String

    Set db = getdb()
    If Not TablaExiste(db, "tbLogErrores") Then GoTo Fin

    sql = "INSERT INTO tbLogErrores (fechaHora, usuario, suplantadoPor, modulo, procedimiento, numeroError, descripcionError, contexto) " & _
          "VALUES (Now(), Environ('USERNAME'), Null, " & _
          "'" & ReemplazarComillas(origen) & "', " & _
          "'" & ReemplazarComillas(procedimiento) & "', " & numero & ", " & _
          "'" & ReemplazarComillas(descripcion) & "', " & _
          "'" & ReemplazarComillas(ConstruirContexto()) & "');"
    db.Execute sql, dbFailOnError
Fin:
    On Error Resume Next
End Sub

Private Function TablaExiste(ByVal db As DAO.Database, ByVal nombreTabla As String) As Boolean
    On Error GoTo NoExiste
    Dim tdf As DAO.TableDef
    Set tdf = db.TableDefs(nombreTabla)
    TablaExiste = True
    Exit Function
NoExiste:
    TablaExiste = False
End Function

Private Function ReemplazarComillas(ByVal s As String) As String
    ReemplazarComillas = Replace(Nz(s, ""), "'", "''")
End Function

Private Function ConstruirContexto() As String
    Dim frm As Access.Form
    On Error Resume Next
    Set frm = Screen.ActiveForm
    If frm Is Nothing Then
        ConstruirContexto = "Sin formulario activo"
    Else
        ConstruirContexto = "Formulario activo: " & frm.name
    End If
End Function
Public Function RegistroExiste(ByVal tabla As String, ByVal campoID As String, ByVal valorID As Long, ByRef db As DAO.Database) As Boolean
    Dim rcd As DAO.Recordset
    Dim sql As String
    On Error GoTo Errores
    
    sql = "SELECT *  FROM [" & tabla & "] WHERE [" & campoID & "] = " & valorID & ";"
    Set rcd = db.OpenRecordset(sql, dbOpenSnapshot)
    RegistroExiste = Not rcd.EOF

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
   
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "FUNCIONES UTILES.RegistroExiste"
    errObj.AddToCallStack "Tabla: " & tabla & ", Campo: " & campoID & ", Valor: " & valorID
    errObj.Raise
End Function
Public Sub CommitEditsOnForm(ByVal frm As Access.Form)
    On Error GoTo Fin
    If frm Is Nothing Then Exit Sub
    
    ' Si hay un control bloqueante (ej. botón), saltamos a un control editable
    Dim c As Control
    For Each c In frm.Controls
        If (c.ControlType = acTextBox Or c.ControlType = acComboBox Or c.ControlType = acCheckBox) Then
            If c.Enabled And Not c.Locked And c.Visible Then
                c.SetFocus
                Exit For
            End If
        End If
    Next c
    
    If frm.Dirty Then frm.Dirty = False
Fin:
    On Error Resume Next
End Sub
Public Sub CommitEditsOnActiveSubform(ByVal parentForm As Access.Form)
    Dim ctlBtn As Access.Control
    Dim subFrm As Access.Form
    Dim ok As Boolean
    
    On Error GoTo SalirSeguro  ' Nunca debe romper el flujo
    
    ' 1) Recuerda el control activo actual (suele ser el botón cmdGrabar)
    Set ctlBtn = Screen.ActiveControl
    
    ' 2) Obtiene el formulario cargado dentro del NavigationSubform
    If parentForm Is Nothing Then GoTo SalirSeguro
    If parentForm.NavigationSubform Is Nothing Then GoTo SalirSeguro
    If parentForm.NavigationSubform.Form Is Nothing Then GoTo SalirSeguro
    Set subFrm = parentForm.NavigationSubform.Form
    
    ' 3) Intenta enfocar cualquier control "focable" del subformulario
    ok = TryFocusAnyControl(subFrm)
    
    ' 4) Devuelve el foco al botón (o al control que estuviera)
    If Not ctlBtn Is Nothing Then
        On Error Resume Next
        ctlBtn.SetFocus
        On Error GoTo 0
    End If
    
SalirSeguro:
    On Error Resume Next
End Sub

' Busca el primer control del formulario que pueda recibir foco
' (Visible, Enabled y que no sea etiqueta). Devuelve True si logró hacer SetFocus.
Public Function TryFocusAnyControl(ByVal frm As Access.Form) As Boolean
    Dim c As Access.Control
    On Error GoTo Fallo
    For Each c In frm.Controls
        If c.Visible Then
            ' Muchos controles permiten foco incluso Locked=True si Enabled=True
            If ControlPuedeRecibirFoco(c) Then
                On Error Resume Next
                c.SetFocus
                If Err.Number = 0 Then
                    TryFocusAnyControl = True
                    Exit Function
                End If
                Err.Clear
                On Error GoTo Fallo
            End If
        End If
    Next c
    Exit Function
Fallo:
    ' Si algún control da error al consultar propiedades, seguimos con el siguiente
    Resume Next
End Function

' Heurística simple para saber si un control puede recibir foco
Private Function ControlPuedeRecibirFoco(ByVal c As Access.Control) As Boolean
    Select Case c.ControlType
        Case acTextBox, acComboBox, acListBox, acCheckBox, acOptionGroup, acToggleButton, acOptionButton, acSubform
            ControlPuedeRecibirFoco = (c.Enabled = True)
        Case Else
            ControlPuedeRecibirFoco = False
    End Select
End Function
Public Sub AplicarEstiloBloqueo(ByRef ctl As Control, ByVal bloqueado As Boolean)
    On Error Resume Next
    If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
        ctl.Locked = bloqueado
        ctl.Enabled = True
        ctl.BorderStyle = 1
        If bloqueado Then
            ctl.BackColor = 15921906
            ctl.BorderColor = 12419407
        Else
            ctl.BackColor = 16777215
            ctl.BorderColor = 12419407
        End If
    End If
End Sub
Public Sub HabilitarControlesSubformulario(ByVal frmSub As Form, ByVal habilitar As Boolean)
    ' RESPONSABILIDAD: Recorre todos los controles de un formulario y los habilita/deshabilita.
    '                  Omite etiquetas, líneas y el botón de guardar, que se gestiona aparte.
    Dim ctl As Control
    On Error Resume Next ' Si un control no tiene la propiedad .Enabled, lo ignoramos

    For Each ctl In frmSub.Controls
        ' --- ESTA ES LA LÓGICA CLAVE ---
        If InStr(1, ctl.Tag, "INMUTABLE") <> 0 Then
            GoTo siguiente ' Salta al siguiente control si la etiqueta contiene "INMUTABLE"
        End If
        ' --- FIN DE LA LÓGICA CLAVE ---
        
        Select Case ctl.ControlType
            Case acLabel, acLine, acRectangle, acImage
                ' Ignorar controles pasivos
            Case acCommandButton
                ' Ignorar botones
            Case acTextBox, acComboBox
                If ctl.name <> "cmdGuardar" Then
                    AplicarEstiloBloqueo ctl, Not habilitar
                End If
            Case acCheckBox
                If ctl.name <> "cmdGuardar" Then
                    ctl.Locked = Not habilitar
                    ctl.Enabled = True
                End If
            Case Else
                If ctl.name <> "cmdGuardar" Then
                    ctl.Enabled = habilitar
                End If
        End Select
siguiente:
    Next ctl
End Sub
Public Function CalcularDuracion(ByVal fechaInicio As Date, ByVal fechaFin As Date) As String
    ' ... (el código de la función es exactamente el mismo, solo cambia a Public) ...
    Dim totalMinutos As Long
    Dim dias As Long, horas As Long, minutos As Long
    
    If fechaFin <= fechaInicio Then
        CalcularDuracion = "0m" ' Devolvemos "0m" en lugar de "N/A" para consistencia
        Exit Function
    End If
    
    totalMinutos = DateDiff("n", fechaInicio, fechaFin)
    
    dias = totalMinutos \ 1440
    horas = (totalMinutos Mod 1440) \ 60
    minutos = (totalMinutos Mod 1440) Mod 60
    
    Dim resultado As String
    If dias > 0 Then resultado = dias & "d "
    If horas > 0 Then resultado = resultado & horas & "h "
    resultado = resultado & minutos & "m"
    
    CalcularDuracion = Trim(resultado)
End Function

Public Function JoinCollection(col As Collection, delimiter As String) As String
    Dim result As String
    Dim i As Long
    
    If col.count = 0 Then
        JoinCollection = ""
        Exit Function
    End If
    
    result = ""
    For i = 1 To col.count
        result = result & col(i)
        If i < col.count Then
            result = result & delimiter
        End If
    Next i
    
    JoinCollection = result
End Function
Public Sub GestionarRibbon(ByVal mostrar As Boolean)
    ' RESPONSABILIDAD: Muestra u oculta la cinta de opciones de Access.
    On Error Resume Next ' Si hay algún problema, no debe detener el arranque.
    
    If mostrar Then
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
    Else
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
    End If
End Sub



