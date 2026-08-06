Attribute VB_Name = "Test_BackendConfigPaths"
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Public Function Test_BackendConfigPaths_SanitizeDifferentUserIsReplaced() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: ruta de otro usuario y perfil actual inyectado"
    logs(1) = "2. Act: SanitizarRutaUsuarioWindowsLocal"
    logs(2) = "3. Assert: reemplaza solo el prefijo C:\\Users\\<usuario>"
    logs(3) = "4. Assert: conserva el resto de la ruta"

    Dim sut As String
    sut = SanitizarRutaUsuarioWindowsLocal("C:\Users\adm1\App\GestionRiesgos\", "C:\Users\currentUser")
    If StrComp(sut, "C:\Users\currentUser\App\GestionRiesgos\", vbTextCompare) <> 0 Then
        Test_BackendConfigPaths_SanitizeDifferentUserIsReplaced = BuildFail("Debe reemplazar el usuario original por el perfil actual inyectado. Actual=[" & sut & "]", logs)
        Exit Function
    End If

    Test_BackendConfigPaths_SanitizeDifferentUserIsReplaced = BuildOk("sanitize_different_user_replaced", logs)
    Exit Function

EH:
    Test_BackendConfigPaths_SanitizeDifferentUserIsReplaced = BuildFail("Test_BackendConfigPaths_SanitizeDifferentUserIsReplaced: " & Err.Description, logs)
End Function

Public Function Test_BackendConfigPaths_SanitizeSameUserIsUnchanged() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: ruta ya apunta al perfil actual"
    logs(1) = "2. Act: SanitizarRutaUsuarioWindowsLocal"
    logs(2) = "3. Assert: mantiene exactamente la misma ruta"

    Dim expectedValue As String
    Dim actual As String
    expectedValue = "C:\Users\currentUser\App\GestionRiesgos\"
    actual = SanitizarRutaUsuarioWindowsLocal(expectedValue, "C:\Users\currentUser")

    If StrComp(actual, expectedValue, vbTextCompare) <> 0 Then
        Test_BackendConfigPaths_SanitizeSameUserIsUnchanged = BuildFail("La ruta con perfil actual debe quedar sin cambios", logs)
        Exit Function
    End If

    Test_BackendConfigPaths_SanitizeSameUserIsUnchanged = BuildOk("sanitize_same_user_unchanged", logs)
    Exit Function

EH:
    Test_BackendConfigPaths_SanitizeSameUserIsUnchanged = BuildFail("Test_BackendConfigPaths_SanitizeSameUserIsUnchanged: " & Err.Description, logs)
End Function

Public Function Test_BackendConfigPaths_SanitizeNonUsersPathIsUnchanged() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: ruta fuera de C:\\Users"
    logs(1) = "2. Act: SanitizarRutaUsuarioWindowsLocal"
    logs(2) = "3. Assert: no se reescribe el prefijo"

    Dim expectedValue As String
    Dim actual As String
    expectedValue = "D:\Apps\GestionRiesgos\"
    actual = SanitizarRutaUsuarioWindowsLocal(expectedValue, "C:\Users\currentUser")

    If StrComp(actual, expectedValue, vbBinaryCompare) <> 0 Then
        Test_BackendConfigPaths_SanitizeNonUsersPathIsUnchanged = BuildFail("Las rutas fuera de C:\\Users no deben mutar", logs)
        Exit Function
    End If

    Test_BackendConfigPaths_SanitizeNonUsersPathIsUnchanged = BuildOk("sanitize_non_users_unchanged", logs)
    Exit Function

EH:
    Test_BackendConfigPaths_SanitizeNonUsersPathIsUnchanged = BuildFail("Test_BackendConfigPaths_SanitizeNonUsersPathIsUnchanged: " & Err.Description, logs)
End Function

Public Function Test_BackendConfigPaths_SanitizeUNCPathIsUnchanged() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: ruta UNC"
    logs(1) = "2. Act: SanitizarRutaUsuarioWindowsLocal"
    logs(2) = "3. Assert: no debe tocar rutas UNC"

    Dim expectedValue As String
    Dim actual As String
    expectedValue = "\\fileserver\Apps\GestionRiesgos\"
    actual = SanitizarRutaUsuarioWindowsLocal(expectedValue, "C:\Users\currentUser")

    If StrComp(actual, expectedValue, vbBinaryCompare) <> 0 Then
        Test_BackendConfigPaths_SanitizeUNCPathIsUnchanged = BuildFail("La ruta UNC no debe ser reescrita", logs)
        Exit Function
    End If

    Test_BackendConfigPaths_SanitizeUNCPathIsUnchanged = BuildOk("sanitize_unc_unchanged", logs)
    Exit Function

EH:
    Test_BackendConfigPaths_SanitizeUNCPathIsUnchanged = BuildFail("Test_BackendConfigPaths_SanitizeUNCPathIsUnchanged: " & Err.Description, logs)
End Function

Public Function Test_BackendConfigPaths_SanitizeEmptyPathSafe() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: ruta vacia"
    logs(1) = "2. Act: SanitizarRutaUsuarioWindowsLocal"
    logs(2) = "3. Assert: resultado vacío y sin error"

    Dim actual As String
    actual = SanitizarRutaUsuarioWindowsLocal("", "C:\Users\currentUser")
    If actual <> "" Then
        Test_BackendConfigPaths_SanitizeEmptyPathSafe = BuildFail("Una ruta vacía debe permanecer vacía", logs)
        Exit Function
    End If

    Test_BackendConfigPaths_SanitizeEmptyPathSafe = BuildOk("sanitize_empty_path_safe", logs)
    Exit Function

EH:
    Test_BackendConfigPaths_SanitizeEmptyPathSafe = BuildFail("Test_BackendConfigPaths_SanitizeEmptyPathSafe: " & Err.Description, logs)
End Function

Public Function Test_BackendConfigPaths_SanitizeKeepsTrailingSeparatorBehavior() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: mismo usuario con o sin barra al final"
    logs(1) = "2. Act: SanitizarRutaUsuarioWindowsLocal con usuario distinto"
    logs(2) = "3. Assert: conserva la forma del separador/sufijo"
    logs(3) = "4. Assert: caso A mantiene slash final"
    logs(4) = "5. Assert: caso B conserva ausencia de slash final"

    Dim withSlash As String
    Dim withoutSlash As String
    withSlash = SanitizarRutaUsuarioWindowsLocal("C:\Users\adm1\App\GestionRiesgos\", "C:\Users\currentUser")
    withoutSlash = SanitizarRutaUsuarioWindowsLocal("C:\Users\adm1\App\GestionRiesgos", "C:\Users\currentUser")

    If StrComp(withSlash, "C:\Users\currentUser\App\GestionRiesgos\", vbTextCompare) <> 0 Then
        Test_BackendConfigPaths_SanitizeKeepsTrailingSeparatorBehavior = BuildFail("La variante con barra final debe conservarla. Actual=[" & withSlash & "]", logs)
        Exit Function
    End If

    If StrComp(withoutSlash, "C:\Users\currentUser\App\GestionRiesgos", vbTextCompare) <> 0 Then
        Test_BackendConfigPaths_SanitizeKeepsTrailingSeparatorBehavior = BuildFail("La variante sin barra final debe conservarla sin barra", logs)
        Exit Function
    End If

    Test_BackendConfigPaths_SanitizeKeepsTrailingSeparatorBehavior = BuildOk("sanitize_separator_behavior", logs)
    Exit Function

EH:
    Test_BackendConfigPaths_SanitizeKeepsTrailingSeparatorBehavior = BuildFail("Test_BackendConfigPaths_SanitizeKeepsTrailingSeparatorBehavior: " & Err.Description, logs)
End Function

Public Function Test_BackendConfigPaths_LeeConfiguracionLocal_SanitizesLocalUserPaths() As String
    On Error GoTo EH

    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: ResetGlobals para limpiar estado previo"
    logs(1) = "2. Act: LeeConfiguracionLocal lee TbConfiguracionBackends"
    logs(2) = "3. Assert: si RutaDirApp_LOCAL es C:\Users, usa USERPROFILE actual"
    logs(3) = "4. Assert: URLRutaAplicacionesLocal queda bajo el perfil actual"
    logs(4) = "5. Assert: BackendSandboxURL queda bajo el perfil actual si es local"
    logs(5) = "6. Cleanup: no modifica datos persistidos"

    Dim errText As String
    Dim currentProfile As String
    Dim currentProfileWithSlash As String

    currentProfile = Environ$("USERPROFILE")
    If Len(currentProfile) = 0 Then
        Test_BackendConfigPaths_LeeConfiguracionLocal_SanitizesLocalUserPaths = BuildFail("USERPROFILE no está disponible para validar la sanitización", logs)
        Exit Function
    End If

    currentProfileWithSlash = currentProfile
    If Right$(currentProfileWithSlash, 1) <> "\" Then
        currentProfileWithSlash = currentProfileWithSlash & "\"
    End If

    ResetGlobals errText
    If Len(errText) > 0 Then
        Test_BackendConfigPaths_LeeConfiguracionLocal_SanitizesLocalUserPaths = BuildFail("ResetGlobals falló: " & errText, logs)
        Exit Function
    End If

    LeeConfiguracionLocal errText
    If Len(errText) > 0 Then
        Test_BackendConfigPaths_LeeConfiguracionLocal_SanitizesLocalUserPaths = BuildFail("LeeConfiguracionLocal falló: " & errText, logs)
        Exit Function
    End If

    If Len(m_RutaDirApp_LOCAL) > 0 Then
        If StrComp(Left$(m_RutaDirApp_LOCAL, Len("C:\Users\")), "C:\Users\", vbTextCompare) = 0 Then
            If StrComp(Left$(m_RutaDirApp_LOCAL, Len(currentProfileWithSlash)), currentProfileWithSlash, vbTextCompare) <> 0 Then
                Test_BackendConfigPaths_LeeConfiguracionLocal_SanitizesLocalUserPaths = BuildFail("m_RutaDirApp_LOCAL no usa USERPROFILE actual. Actual=[" & m_RutaDirApp_LOCAL & "] Perfil=[" & currentProfileWithSlash & "]", logs)
                Exit Function
            End If
        End If
    End If

    If Len(m_URLRutaAplicacionesLocal) > 0 Then
        If StrComp(Left$(m_URLRutaAplicacionesLocal, Len("C:\Users\")), "C:\Users\", vbTextCompare) = 0 Then
            If StrComp(Left$(m_URLRutaAplicacionesLocal, Len(currentProfileWithSlash)), currentProfileWithSlash, vbTextCompare) <> 0 Then
                Test_BackendConfigPaths_LeeConfiguracionLocal_SanitizesLocalUserPaths = BuildFail("m_URLRutaAplicacionesLocal no usa USERPROFILE actual. Actual=[" & m_URLRutaAplicacionesLocal & "] Perfil=[" & currentProfileWithSlash & "]", logs)
                Exit Function
            End If
        End If
    End If

    If Len(m_BackendSandboxURL) > 0 Then
        If StrComp(Left$(m_BackendSandboxURL, Len("C:\Users\")), "C:\Users\", vbTextCompare) = 0 Then
            If StrComp(Left$(m_BackendSandboxURL, Len(currentProfileWithSlash)), currentProfileWithSlash, vbTextCompare) <> 0 Then
                Test_BackendConfigPaths_LeeConfiguracionLocal_SanitizesLocalUserPaths = BuildFail("m_BackendSandboxURL no usa USERPROFILE actual. Actual=[" & m_BackendSandboxURL & "] Perfil=[" & currentProfileWithSlash & "]", logs)
                Exit Function
            End If
        End If
    End If

    Test_BackendConfigPaths_LeeConfiguracionLocal_SanitizesLocalUserPaths = BuildOk("lee_configuracion_local_sanitizes_user_paths", logs)
    Exit Function

EH:
    Test_BackendConfigPaths_LeeConfiguracionLocal_SanitizesLocalUserPaths = BuildFail("Test_BackendConfigPaths_LeeConfiguracionLocal_SanitizesLocalUserPaths: " & Err.Description, logs)
End Function
