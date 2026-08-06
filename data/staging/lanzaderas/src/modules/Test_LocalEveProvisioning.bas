Attribute VB_Name = "Test_LocalEveProvisioning"
Option Compare Database
Option Explicit

Private Function NormalizeDirectoryPath(ByVal p_Path As String) As String
    Dim normalizedPath As String
    normalizedPath = Trim$(Replace(p_Path, "/", "\"))
    If normalizedPath <> "" And Right$(normalizedPath, 1) <> "\" Then normalizedPath = normalizedPath & "\"
    NormalizeDirectoryPath = normalizedPath
End Function

Private Function AppendLocalPath(ByVal p_BasePath As String, ByVal p_ChildPath As String) As String
    AppendLocalPath = NormalizeDirectoryPath(p_BasePath) & p_ChildPath
End Function

Private Function LastDirectoryName(ByVal p_Path As String) As String
    Dim normalizedPath As String
    Dim pathWithoutTrailingSlash As String
    Dim pathParts As Variant

    normalizedPath = NormalizeDirectoryPath(p_Path)
    If normalizedPath = "" Then Exit Function

    pathWithoutTrailingSlash = Left$(normalizedPath, Len(normalizedPath) - 1)
    pathParts = Split(pathWithoutTrailingSlash, "\")
    LastDirectoryName = CStr(pathParts(UBound(pathParts)))
End Function

Private Function ParentDirectoryPath(ByVal p_Path As String) As String
    Dim normalizedPath As String
    Dim pathWithoutTrailingSlash As String
    Dim lastSeparatorPos As Long

    normalizedPath = NormalizeDirectoryPath(p_Path)
    If normalizedPath = "" Then Exit Function

    pathWithoutTrailingSlash = Left$(normalizedPath, Len(normalizedPath) - 1)
    lastSeparatorPos = InStrRev(pathWithoutTrailingSlash, "\")
    If lastSeparatorPos = 0 Then Exit Function

    ParentDirectoryPath = Left$(pathWithoutTrailingSlash, lastSeparatorPos)
End Function

Private Sub ResolveConfiguredApplicationPaths(ByVal p_ConfiguredPath As String, _
                                             ByRef p_AppRoot As String, _
                                             ByRef p_AppFolder As String)
    Dim normalizedPath As String

    normalizedPath = NormalizeDirectoryPath(p_ConfiguredPath)
    If StrComp(LastDirectoryName(normalizedPath), "0Lanzadera", vbTextCompare) = 0 Then
        p_AppFolder = normalizedPath
        p_AppRoot = ParentDirectoryPath(normalizedPath)
    Else
        p_AppRoot = normalizedPath
        p_AppFolder = AppendLocalPath(p_AppRoot, "0Lanzadera\")
    End If
End Sub

Private Function IsLocalDrivePath(ByVal p_Path As String) As Boolean
    Dim normalizedPath As String
    normalizedPath = Trim$(p_Path)
    If normalizedPath = "" Then Exit Function
    If Left$(normalizedPath, 2) = "\\" Then Exit Function
    IsLocalDrivePath = (InStr(1, normalizedPath, ":\", vbTextCompare) > 0)
End Function

Private Sub AddMissingPath(ByRef p_Missing As String, _
                           ByVal p_Name As String, _
                           ByVal p_Path As String)
    If p_Missing <> "" Then p_Missing = p_Missing & "; "
    p_Missing = p_Missing & p_Name & "=" & p_Path
End Sub

Private Sub RequireLocalPath(ByRef p_Missing As String, _
                              ByVal p_Name As String, _
                              ByVal p_Path As String, _
                             ByVal p_IsFile As Boolean)
    Dim localFso As Object
    Set localFso = CreateObject("Scripting.FileSystemObject")

    If p_IsFile Then
        If Not localFso.FileExists(p_Path) Then AddMissingPath p_Missing, p_Name, p_Path
    Else
        If Not localFso.FolderExists(p_Path) Then AddMissingPath p_Missing, p_Name, p_Path
    End If
End Sub

Private Sub RequireAnyLocalFile(ByRef p_Missing As String, _
                                ByVal p_Name As String, _
                                ByVal p_PrimaryPath As String, _
                                ByVal p_LegacyPath As String)
    Dim localFso As Object
    Set localFso = CreateObject("Scripting.FileSystemObject")

    If localFso.FileExists(p_PrimaryPath) Then Exit Sub
    If localFso.FileExists(p_LegacyPath) Then Exit Sub

    AddMissingPath p_Missing, p_Name, p_PrimaryPath & " OR " & p_LegacyPath
End Sub

Private Function ValidateLocalEveResources(Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim backendActivo As String
    Dim backendSandbox As String
    Dim configuredAppPath As String
    Dim appRoot As String
    Dim appFolder As String
    Dim resourcesFolder As String
    Dim missingPaths As String

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT TOP 2 BackendActivo, BackendSandbox, RutaDirectorioAplicacion_LOCAL FROM TbConfiguracionBackends", dbOpenSnapshot)
    If rs.EOF Then
        p_Error = "TESTS BLOCKED: TbConfiguracionBackends no tiene configuración local para EVE"
        GoTo Cleanup
    End If

    backendActivo = UCase$(Trim$(Nz(rs.Fields("BackendActivo").value, "")))
    backendSandbox = Trim$(Nz(rs.Fields("BackendSandbox").value, ""))
    configuredAppPath = NormalizeDirectoryPath(CStr(Nz(rs.Fields("RutaDirectorioAplicacion_LOCAL").value, "")))
    ResolveConfiguredApplicationPaths configuredAppPath, appRoot, appFolder
    resourcesFolder = AppendLocalPath(appFolder, "recursos\")

    rs.MoveNext
    If Not rs.EOF Then
        p_Error = "TESTS BLOCKED: TbConfiguracionBackends debe tener una sola fila para validar EVE local"
        GoTo Cleanup
    End If

    If backendActivo <> "LOCAL" Then
        p_Error = "TESTS BLOCKED: BackendActivo debe ser LOCAL para validar EVE; actual=" & backendActivo
        GoTo Cleanup
    End If
    If Not IsLocalDrivePath(backendSandbox) Then
        p_Error = "TESTS BLOCKED: BackendSandbox no es una ruta local segura: " & backendSandbox
        GoTo Cleanup
    End If
    If Not IsLocalDrivePath(configuredAppPath) Then
        p_Error = "TESTS BLOCKED: RutaDirectorioAplicacion_LOCAL no es una ruta local segura: " & configuredAppPath
        GoTo Cleanup
    End If
    If Not IsLocalDrivePath(appRoot) Then
        p_Error = "TESTS BLOCKED: no se pudo resolver la raíz de aplicaciones desde RutaDirectorioAplicacion_LOCAL: " & configuredAppPath
        GoTo Cleanup
    End If

    RequireLocalPath missingPaths, "BackendSandbox", backendSandbox, True
    RequireLocalPath missingPaths, "RutaDirectorioAplicacion_LOCAL", configuredAppPath, False
    RequireLocalPath missingPaths, "URLDirAplicaciones", appRoot, False
    RequireLocalPath missingPaths, "URLDirAplicacionLanzadera", appFolder, False
    RequireLocalPath missingPaths, "URLDirRecursosLanzadera", resourcesFolder, False
    RequireLocalPath missingPaths, "URLDirAyuda", AppendLocalPath(appFolder, "ANEXOS\AYUDA\"), False
    RequireLocalPath missingPaths, "URLDirIconosArbol", AppendLocalPath(resourcesFolder, "IconosArbol\"), False
    RequireLocalPath missingPaths, "URLArchivoCSS", AppendLocalPath(appRoot, "css.txt"), True
    RequireAnyLocalFile missingPaths, "URLAyudaInstalarNetFrameWork", AppendLocalPath(resourcesFolder, "InstalarFrx.pdf"), AppendLocalPath(appRoot, "InstalarFrx.pdf")
    RequireLocalPath missingPaths, "URLArchivoIniLanzadera", AppendLocalPath(resourcesFolder, "Lanzadera.ini"), True

    If missingPaths <> "" Then
        p_Error = "TESTS BLOCKED: faltan recursos locales requeridos por EVE: " & missingPaths
        GoTo Cleanup
    End If

    ValidateLocalEveResources = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
EH:
    p_Error = "ValidateLocalEveResources: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

Private Function ErrorMentionsCriticalEntornoProperty(ByVal p_Error As String) As Boolean
    Dim propertyNames As Variant
    Dim i As Long
    propertyNames = Array("URLDirAplicaciones", "URLDirAplicacionLanzadera", "URLDirAyuda", "URLDirRecursosLanzadera", "CSS", "URLArchivoCSS", "URLArchivoIniLanzadera", "URLAyudaInstalarNetFrameWork", "VersionAplicacion", "URLDirIconosArbol", "URLDirVideos", "CabeceraHTML", "URLConfiguracion", "PermitirDiaSinContrasenia", "ReiniciarVersiones")

    For i = LBound(propertyNames) To UBound(propertyNames)
        If InStr(1, p_Error, CStr(propertyNames(i)), vbTextCompare) > 0 Then
            ErrorMentionsCriticalEntornoProperty = True
            Exit Function
        End If
    Next i
End Function

Public Function Test_LocalEveProvisioning_EVE_NoReportaRecursosFaltantes() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim errMsg As String
    Dim cleanupError As String
    Dim resultJson As String

    logs(0) = "1. Arrange: validar recursos locales requeridos por EVE"
    If Not ValidateLocalEveResources(errMsg) Then
        Test_LocalEveProvisioning_EVE_NoReportaRecursosFaltantes = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: activar sandbox local"
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_LocalEveProvisioning_EVE_NoReportaRecursosFaltantes = Test_Helper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    logs(2) = "3. Act: ejecutar EVE forzado a LOCAL"
    errMsg = ""
    EVE errMsg, "LOCAL"

    logs(3) = "4. Assert: EVE no reporta propiedades Entorno faltantes"
    If ErrorMentionsCriticalEntornoProperty(errMsg) Then
        resultJson = Test_Helper.BuildJsonFail("EVE reportó propiedades Entorno faltantes con recursos locales presentes: " & errMsg, logs)
        GoTo Cleanup
    End If
    If errMsg <> "" Then
        resultJson = Test_Helper.BuildJsonFail("EVE falló en LOCAL con error no esperado: " & errMsg, logs)
        GoTo Cleanup
    End If

    resultJson = Test_Helper.BuildJsonOk("local_eve_resources_ok", logs)

Cleanup:
    logs(4) = "5. Cleanup: ResetTestSession"
    If Not Test_Helper.ResetTestSession(cleanupError) Then
        resultJson = Test_Helper.BuildJsonFail(cleanupError, logs)
    End If
    Test_LocalEveProvisioning_EVE_NoReportaRecursosFaltantes = resultJson
    Exit Function
EH:
    resultJson = Test_Helper.BuildJsonFail(Err.Description, logs)
    Resume Cleanup
End Function
