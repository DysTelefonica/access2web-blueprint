Attribute VB_Name = "Test_InfraestructuraInicio"
Option Compare Database
Option Explicit

Private Function JsonOk(ByVal p_Value As String) As String
    Dim logs(0 To 0) As String
    logs(0) = "ok: " & p_Value
    JsonOk = BuildJsonOk(p_Value, logs)
End Function

Private Function JsonFail(ByVal p_Error As String) As String
    Dim logs(0 To 0) As String
    logs(0) = "fail: " & p_Error
    JsonFail = BuildJsonFail(p_Error, logs)
End Function

Public Function Test_ValidarCarpetaEscribible_DetectaRutaNoAlcanzable() As String
    Dim m_Error As String
    Dim m_RutaInexistente As String

    ' Default: fail. Solo se override si los asserts pasan.
    Test_ValidarCarpetaEscribible_DetectaRutaNoAlcanzable = JsonFail("test did not complete")

    On Error GoTo HandleError
    m_RutaInexistente = Environ$("TEMP") & "\EXPEDIENTES_TEST_NO_EXISTE_" & Format$(Now, "yyyymmddhhnnss") & "\"

    ValidarCarpetaEscribible m_RutaInexistente, "documentación de anexos", m_Error
    If m_Error = "" Then
        Test_ValidarCarpetaEscribible_DetectaRutaNoAlcanzable = JsonFail("La validación aceptó una ruta inexistente")
        Exit Function
    End If
    If InStr(1, m_Error, "No es alcanzable la ruta de documentación de anexos", vbTextCompare) = 0 Then
        Test_ValidarCarpetaEscribible_DetectaRutaNoAlcanzable = JsonFail("Mensaje inesperado: " & m_Error)
        Exit Function
    End If

    Test_ValidarCarpetaEscribible_DetectaRutaNoAlcanzable = JsonOk("missing_path_detected")
    Exit Function
HandleError:
    Test_ValidarCarpetaEscribible_DetectaRutaNoAlcanzable = JsonFail(Err.Description)
End Function

Public Function Test_ValidarCarpetaEscribible_AceptaRutaTemporal() As String
    Dim m_Error As String
    Dim m_RutaTemporal As String

    ' Default: fail. Solo se override si los asserts pasan.
    Test_ValidarCarpetaEscribible_AceptaRutaTemporal = JsonFail("test did not complete")

    On Error GoTo HandleError
    m_RutaTemporal = Environ$("TEMP") & "\EXPEDIENTES_TEST_INFRA_" & Format$(Now, "yyyymmddhhnnss") & "\"
    fso.CreateFolder m_RutaTemporal

    ValidarCarpetaEscribible m_RutaTemporal, "documentación de anexos", m_Error
    If m_Error <> "" Then
        Test_ValidarCarpetaEscribible_AceptaRutaTemporal = JsonFail(m_Error)
        GoTo Cleanup
    End If

    Test_ValidarCarpetaEscribible_AceptaRutaTemporal = JsonOk("temp_path_writable")
Cleanup:
    On Error Resume Next
    ' Cleanup best-effort: borra archivos dentro + carpeta. Cero dialog
    ' (RmDir + Kill nativos, no fso.DeleteFolder que puede tirar dialogs).
    If Len(m_RutaTemporal) > 0 And Dir(m_RutaTemporal, vbDirectory) <> "" Then
        Dim m_F As String
        m_F = Dir(m_RutaTemporal & "\*.*")
        Do While Len(m_F) > 0
            Kill m_RutaTemporal & "\" & m_F
            m_F = Dir()
        Loop
        RmDir StripTrailingSlash(m_RutaTemporal)
    End If
    Exit Function
HandleError:
    Test_ValidarCarpetaEscribible_AceptaRutaTemporal = JsonFail(Err.Description)
    Resume Cleanup
End Function

Public Function Test_BackendConfigCompat_LeeConfiguracionLocal_InvalidKeyFailsFastAndCleansState() As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_BackendConfigCompat_LeeConfiguracionLocal_InvalidKeyFailsFastAndCleansState = JsonFail("test did not complete")

    On Error GoTo HandleError

    Dim m_Error As String
    Dim m_Db As DAO.Database
    Dim m_DbError As String
    Dim m_ProdPath As String
    Dim m_Handle As Integer

    m_ProdPath = Environ$("TEMP") & "\EXPEDIENTES_TEST_BACKEND_COMPAT_PROD_" & Format$(Now, "yyyymmddhhnnss") & ".accdb"
    m_Handle = FreeFile
    Open m_ProdPath For Output As #m_Handle
    Close #m_Handle

    TestOnlyResetBackendConfigOverride
    TestOnlySetBackendConfigOverride "OTRO", m_ProdPath, m_ProdPath, m_ProdPath, "pwd-test"

    LeeConfiguracionLocal m_Error
    If m_Error = "" Then
        Test_BackendConfigCompat_LeeConfiguracionLocal_InvalidKeyFailsFastAndCleansState = JsonFail("Se esperaba error por backendActivo inválido")
        GoTo Cleanup
    End If

    If InStr(1, m_Error, "backendActivo", vbTextCompare) = 0 Then
        Test_BackendConfigCompat_LeeConfiguracionLocal_InvalidKeyFailsFastAndCleansState = JsonFail("Mensaje inesperado: " & m_Error)
        GoTo Cleanup
    End If

    If TestOnlyGetActiveBackendURL() <> "" Then
        Test_BackendConfigCompat_LeeConfiguracionLocal_InvalidKeyFailsFastAndCleansState = JsonFail("m_ActiveBackendURL debe quedar vacío tras fallo")
        GoTo Cleanup
    End If

    If TestOnlyGetPasswordBackend() <> "" Then
        Test_BackendConfigCompat_LeeConfiguracionLocal_InvalidKeyFailsFastAndCleansState = JsonFail("m_PasswordBackend debe quedar vacío tras fallo")
        GoTo Cleanup
    End If

    Set m_Db = getdb(m_DbError)
    If Not m_Db Is Nothing Then
        Test_BackendConfigCompat_LeeConfiguracionLocal_InvalidKeyFailsFastAndCleansState = JsonFail("getdb no debe abrir conexión tras fallo de configuración")
        GoTo Cleanup
    End If

    If m_DbError = "" Then
        Test_BackendConfigCompat_LeeConfiguracionLocal_InvalidKeyFailsFastAndCleansState = JsonFail("getdb debe propagar error tras configuración inválida")
        GoTo Cleanup
    End If

    Test_BackendConfigCompat_LeeConfiguracionLocal_InvalidKeyFailsFastAndCleansState = JsonOk("invalid_key_failfast_cleanup")

Cleanup:
    On Error Resume Next
    Set m_Db = Nothing
    TestOnlyResetBackendConfigOverride
    If Dir$(m_ProdPath) <> "" Then Kill m_ProdPath
    Exit Function

HandleError:
    Test_BackendConfigCompat_LeeConfiguracionLocal_InvalidKeyFailsFastAndCleansState = JsonFail(Err.Description)
    Resume Cleanup
End Function

Public Function Test_BackendConfigCompat_GetDbFailFastWhenResolvedPathEmpty() As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_BackendConfigCompat_GetDbFailFastWhenResolvedPathEmpty = JsonFail("test did not complete")

    On Error GoTo HandleError

    Dim m_Error As String
    Dim m_DbError As String
    Dim m_Db As DAO.Database

    TestOnlyResetBackendConfigOverride
    TestOnlySetBackendConfigOverride "PROD", "   ", "C:\sandbox-no-usar.accdb", "C:\test-no-usar.accdb", "pwd-test"

    LeeConfiguracionLocal m_Error
    If m_Error = "" Then
        Test_BackendConfigCompat_GetDbFailFastWhenResolvedPathEmpty = JsonFail("Se esperaba error por ruta backend vacía")
        GoTo Cleanup
    End If

    If TestOnlyGetActiveBackendURL() <> "" Then
        Test_BackendConfigCompat_GetDbFailFastWhenResolvedPathEmpty = JsonFail("m_ActiveBackendURL debe quedar vacío")
        GoTo Cleanup
    End If

    Set m_Db = getdb(m_DbError)
    If Not m_Db Is Nothing Then
        Test_BackendConfigCompat_GetDbFailFastWhenResolvedPathEmpty = JsonFail("getdb no debe abrir DB con ruta vacía")
        GoTo Cleanup
    End If

    If m_DbError = "" Then
        Test_BackendConfigCompat_GetDbFailFastWhenResolvedPathEmpty = JsonFail("getdb debe devolver p_Error cuando la ruta está vacía")
        GoTo Cleanup
    End If

    Test_BackendConfigCompat_GetDbFailFastWhenResolvedPathEmpty = JsonOk("getdb_failfast_empty_path")

Cleanup:
    On Error Resume Next
    Set m_Db = Nothing
    TestOnlyResetBackendConfigOverride
    Exit Function

HandleError:
    Test_BackendConfigCompat_GetDbFailFastWhenResolvedPathEmpty = JsonFail(Err.Description)
    Resume Cleanup
End Function

' Quita el trailing backslash de un path. RmDir no lo acepta.
Private Function StripTrailingSlash(ByVal p_Path As String) As String
    If Len(p_Path) > 0 And Right$(p_Path, 1) = "\" Then
        StripTrailingSlash = Left$(p_Path, Len(p_Path) - 1)
    Else
        StripTrailingSlash = p_Path
    End If
End Function
