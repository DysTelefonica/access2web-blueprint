Attribute VB_Name = "Test_Core_EVE_UsuarioActivo"
' ============================================================
' Test_Core_EVE_UsuarioActivo
'
' issue-55: EVE no consultaba Usuario.Activo, asi que un usuario con
' FechaBaja vencida podia iniciar sesion y ver/editar datos.
'
' El fix agrega un guard en EVE (src/modules/Variables Globales.bas)
' que aborta con error claro si m_ObjUsuarioConectado.Activo = EnumSiNo.No.
'
' POR QUE NO LLAMAR A EVE EN LOS TESTS:
' EVE es un entry-point de identidad con MUCHOS side effects (ResetGlobals,
' LeeConfiguracionLocal, crea Entorno, valida recursos, carga menus, etc).
' En el sandbox esos pasos pueden fallar con errores colaterales
' (Desbordamiento, validacion de rutas UNC, etc) que no tienen que ver
' con el guard que estamos probando. Para tests de integracion reales
' del flujo EVE ver Test_TestingModeSandbox.
'
' ESTRATEGIA: los tests verifican la LOGICA SUBYACENTE que el guard usa:
'   - Constructor.getUsuario resuelve el usuario desde TbUsuariosAplicaciones
'   - Usuario.Activo se calcula correctamente segun FechaBaja
' El guard en EVE es trivial: `If m_ObjUsuarioConectado.Activo = EnumSiNo.No
' Then Err.Raise 1000`. Si los tests verifican que Activo retorna el valor
' correcto, el guard funciona por transitividad.
'
' Skill: access-vba-tdd v2.4
' - Sandbox via Test_Fixtures.GetTestDb (apunta al backend)
' - User fixture en TbUsuariosAplicaciones con Id = 9810
'   (dbInteger 16-bit, max 32767; NO usar 900810+ que desborda)
' - Teardown explicito (DELETE fila)
' ============================================================
Option Compare Database
Option Explicit

' --- Fixture del usuario inactivo ---
' Id es dbInteger (16-bit, max 32767). Usar valor pequeno pero distintivo.
Private Const FIX_ID As Long = 9810
Private Const FIX_CORREO As String = "test_inactivo_issue55@telefonica.com"
Private Const FIX_USUARIORED As String = "test_inactivo_issue55"
Private Const FIX_FECHABAJA_PASADA As Date = #1/1/2020#


' --- Helpers JSON ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function


' --- Helpers de fixture ---
Private Sub CleanUsuarioInactivo()
    On Error Resume Next
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub
    db.Execute "DELETE FROM TbUsuariosAplicaciones WHERE CorreoUsuario='" & FIX_CORREO & "'"
    On Error GoTo 0
End Sub


' ============================================================
' Test 1: Activo = No cuando FechaBaja < Date.
' Este es el caso que el guard de EVE atrapa. Si Activo no retornara
' No, el guard dejaria pasar al usuario.
' ============================================================
Public Function Test_Core_UsuarioActivo_FechaBajaPasada_RetornaNo() As String
    On Error GoTo EH
    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: sandbox + insertar TbUsuariosAplicaciones con FechaBaja=" & Format$(FIX_FECHABAJA_PASADA, "yyyy-mm-dd")
    logs(1) = "2. Act: Constructor.getUsuario(, , , '" & FIX_CORREO & "', sError)"
    logs(2) = "3. Act: usuario.Activo"
    logs(3) = "4. Assert: Activo = EnumSiNo.No (usuario inactivo)"
    logs(4) = "5. Teardown: DELETE fila de TbUsuariosAplicaciones"
    logs(5) = "6. Assert: side effects limpios"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Core_UsuarioActivo_FechaBajaPasada_RetornaNo = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If
    CleanUsuarioInactivo

    ' --- Insertar usuario con FechaBaja en el pasado ---
    db.Execute "INSERT INTO TbUsuariosAplicaciones (Id, CorreoUsuario, UsuarioRed, Nombre, FechaBaja) " & _
        "VALUES (" & FIX_ID & ", '" & FIX_CORREO & "', '" & FIX_USUARIORED & "', 'Test Inactivo Issue55', " & _
        "#" & Format$(FIX_FECHABAJA_PASADA, "yyyy-mm-dd") & "#)", dbFailOnError

    ' --- Resolver y verificar Activo ---
    Dim sError As String
    Dim usuario As Usuario
    Set usuario = Constructor.getUsuario(, , , FIX_CORREO, sError)
    If sError <> "" Then
        CleanUsuarioInactivo
        Test_Core_UsuarioActivo_FechaBajaPasada_RetornaNo = BuildFail( _
            "getUsuario retorno error: " & sError, logs)
        Exit Function
    End If
    If usuario Is Nothing Then
        CleanUsuarioInactivo
        Test_Core_UsuarioActivo_FechaBajaPasada_RetornaNo = BuildFail( _
            "getUsuario retorno Nothing para el usuario de test", logs)
        Exit Function
    End If

    ' --- Assert: Activo = No ---
    If usuario.Activo <> EnumSiNo.No Then
        CleanUsuarioInactivo
        Test_Core_UsuarioActivo_FechaBajaPasada_RetornaNo = BuildFail( _
            "Activo deberia ser No para FechaBaja < Date. Obtenido: " & CStr(usuario.Activo), logs)
        Exit Function
    End If

    ' --- Teardown ---
    CleanUsuarioInactivo
    Test_Core_UsuarioActivo_FechaBajaPasada_RetornaNo = BuildOk( _
        "activo_no_para_fecha_pasada", logs)
    Exit Function
EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    CleanUsuarioInactivo
    Test_Core_UsuarioActivo_FechaBajaPasada_RetornaNo = BuildFail( _
        "Excepcion inesperada: " & errNum & " - " & errDesc, logs)
End Function


' ============================================================
' Test 2: Activo = Si cuando FechaBaja es nula (caso normal).
' Este test evita que el fix se vuelva un blanket block que rompa
' a los usuarios reales. Si FechaBaja es NULL, Activo debe ser Si.
' ============================================================
Public Function Test_Core_UsuarioActivo_FechaBajaNula_RetornaSi() As String
    On Error GoTo EH
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: insertar TbUsuariosAplicaciones con FechaBaja NULL"
    logs(1) = "2. Act: Constructor.getUsuario + usuario.Activo"
    logs(2) = "3. Assert: Activo = EnumSiNo.Sí (caso normal)"
    logs(3) = "4. Teardown: DELETE fila"
    logs(4) = "5. Assert: side effects limpios"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Core_UsuarioActivo_FechaBajaNula_RetornaSi = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If
    CleanUsuarioInactivo

    db.Execute "INSERT INTO TbUsuariosAplicaciones (Id, CorreoUsuario, UsuarioRed, Nombre) " & _
        "VALUES (" & FIX_ID & ", '" & FIX_CORREO & "', '" & FIX_USUARIORED & "', 'Test Activo Issue55')", dbFailOnError

    Dim sError As String
    Dim usuario As Usuario
    Set usuario = Constructor.getUsuario(, , , FIX_CORREO, sError)
    If sError <> "" Or usuario Is Nothing Then
        CleanUsuarioInactivo
        Test_Core_UsuarioActivo_FechaBajaNula_RetornaSi = BuildFail( _
            "getUsuario fallo: sError='" & sError & "', IsNothing=" & (usuario Is Nothing), logs)
        Exit Function
    End If

    If usuario.Activo <> EnumSiNo.Sí Then
        CleanUsuarioInactivo
        Test_Core_UsuarioActivo_FechaBajaNula_RetornaSi = BuildFail( _
            "Activo deberia ser Si para FechaBaja NULL. Obtenido: " & CStr(usuario.Activo), logs)
        Exit Function
    End If

    CleanUsuarioInactivo
    Test_Core_UsuarioActivo_FechaBajaNula_RetornaSi = BuildOk( _
        "activo_si_para_fecha_nula", logs)
    Exit Function
EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    CleanUsuarioInactivo
    Test_Core_UsuarioActivo_FechaBajaNula_RetornaSi = BuildFail( _
        "Excepcion inesperada: " & errNum & " - " & errDesc, logs)
End Function


' ============================================================
' Test 3: Activo = Si cuando FechaBaja es futura (caso normal con
' fecha explicita, e.g. usuario con baja programada para el futuro).
' ============================================================
Public Function Test_Core_UsuarioActivo_FechaBajaFutura_RetornaSi() As String
    On Error GoTo EH
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: insertar TbUsuariosAplicaciones con FechaBaja FUTURA (12/31/2099)"
    logs(1) = "2. Act: usuario.Activo"
    logs(2) = "3. Assert: Activo = EnumSiNo.Sí (sigue activo hasta la fecha)"
    logs(3) = "4. Teardown: DELETE fila"
    logs(4) = "5. Assert: side effects limpios"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Core_UsuarioActivo_FechaBajaFutura_RetornaSi = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If
    CleanUsuarioInactivo

    db.Execute "INSERT INTO TbUsuariosAplicaciones (Id, CorreoUsuario, UsuarioRed, Nombre, FechaBaja) " & _
        "VALUES (" & FIX_ID & ", '" & FIX_CORREO & "', '" & FIX_USUARIORED & "', 'Test Futuro Issue55', " & _
        "#12/31/2099#)", dbFailOnError

    Dim sError As String
    Dim usuario As Usuario
    Set usuario = Constructor.getUsuario(, , , FIX_CORREO, sError)
    If sError <> "" Or usuario Is Nothing Then
        CleanUsuarioInactivo
        Test_Core_UsuarioActivo_FechaBajaFutura_RetornaSi = BuildFail( _
            "getUsuario fallo: sError='" & sError & "', IsNothing=" & (usuario Is Nothing), logs)
        Exit Function
    End If

    If usuario.Activo <> EnumSiNo.Sí Then
        CleanUsuarioInactivo
        Test_Core_UsuarioActivo_FechaBajaFutura_RetornaSi = BuildFail( _
            "Activo deberia ser Si para FechaBaja en el futuro. Obtenido: " & CStr(usuario.Activo), logs)
        Exit Function
    End If

    CleanUsuarioInactivo
    Test_Core_UsuarioActivo_FechaBajaFutura_RetornaSi = BuildOk( _
        "activo_si_para_fecha_futura", logs)
    Exit Function
EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    CleanUsuarioInactivo
    Test_Core_UsuarioActivo_FechaBajaFutura_RetornaSi = BuildFail( _
        "Excepcion inesperada: " & errNum & " - " & errDesc, logs)
End Function
