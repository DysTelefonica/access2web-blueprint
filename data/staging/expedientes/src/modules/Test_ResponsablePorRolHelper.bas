Attribute VB_Name = "Test_ResponsablePorRolHelper"
Option Compare Database
Option Explicit

' -----------------------------------------------------------------------------
' Test_ResponsablePorRolHelper
'
' Atomos TDD para `modResponsablePorRolHelper.GetUsuariosPorRol` y
' `modResponsablePorRolService.ExpedienteGeneral_EstablecerCombos`.  Cada
' test crea su propia temp `.accdb` (DDL minimo + seed controlado) y la
' inyecta como `p_db`.  El helper NUNCA toca `CurrentDb()` real en estos
' tests porque siempre recibe un `p_db` no Nothing (verificar que el
' helper respeta el db inyectado es parte del contrato).
'
' Patron: sandbox §5.5 per-test temp `.accdb` (con §3.9 + §5.4 + §5.5).
' JSON contract §2 (runner): `{ok, value, payload, error, logs}`.
'
' DDL minimo en cada temp .accdb (debe espejar el schema real):
'   - TbResponsablesPorRol     (IDResponsablePorRol COUNTER PK, IDUsuario Long, Rol Text,
'                                UNIQUE INDEX (IDUsuario, Rol))
'   - TbUsuariosAplicaciones   (Id SHORT 16-bit, Nombre, CorreoUsuario, UsuarioRed,
'                                Activado, FechaAlta, FechaBaja)
'
' IDs de fixture: rango 30001..32766 dentro del dominio SHORT (16-bit signed Integer,
' max 32767) que es el tipo real de TbUsuariosAplicaciones.Id.  Cada test crea su
' propia temp .accdb (sandbox §5.5), asi que no hay colision con IDs productivos
' del frontend ni del backend.
'
'   Calidad:  30001..30005
'   Seguridad:30006..30008
'   Cross-rol (Calidad + Seguridad): 30011
'   Baja logica (Calidad + FechaBaja): 30010
'   Orphan en TbUsuariosAplicaciones sin rol: 30012
' -----------------------------------------------------------------------------

Private Const TEST_DIR  As String = "C:\Users\adm1\AppData\Local\Temp\test-responsable-por-rol\"
Private Const FIXTURE_ID_BASE As Long = 30000


' -----------------------------------------------------------------------------
' Atomos publicos (entrypoints del runner)
' -----------------------------------------------------------------------------

Public Function Test_ResponsablePorRol_Calidad_DevuelveUsuariosActivosConRol() As String
    Dim db As DAO.Database
    Dim dic As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Path As String
    Dim m_Logs(0 To 4) As String

    On Error GoTo errores
    m_Path = TEST_DIR & "calidad_happy_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = CreateIsolatedTempDbResponsable(m_Path)
    SeedRolesFixture db

    m_Logs(0) = "1. Arrange: temp .accdb con seed (5 Calidad activos, 3 Seguridad activos)"
    Set dic = modResponsablePorRolHelper.GetUsuariosPorRol( _
                    "Calidad", "", db, m_Error)
    If m_Error <> "" Then
        Test_ResponsablePorRol_Calidad_DevuelveUsuariosActivosConRol = BuildFail("unexpected p_Error", m_Logs, m_Error)
        GoTo Teardown
    End If

    m_Logs(1) = "2. Assert: 5 entradas con Id en {30001..30005}"
    If dic Is Nothing Then
        Test_ResponsablePorRol_Calidad_DevuelveUsuariosActivosConRol = BuildFail("dic es Nothing", m_Logs, "")
        GoTo Teardown
    End If
    If dic.Count <> 5 Then
        Test_ResponsablePorRol_Calidad_DevuelveUsuariosActivosConRol = BuildFail( _
            "Count=" & dic.Count & " esperaba 5", m_Logs, "")
        GoTo Teardown
    End If
    If Not dic.Exists(CStr(FIXTURE_ID_BASE + 1)) Or _
       Not dic.Exists(CStr(FIXTURE_ID_BASE + 5)) Then
        Test_ResponsablePorRol_Calidad_DevuelveUsuariosActivosConRol = BuildFail( _
            "faltan IDs esperados en diccionario", m_Logs, "")
        GoTo Teardown
    End If

    Test_ResponsablePorRol_Calidad_DevuelveUsuariosActivosConRol = BuildOk("ok", m_Logs)
    GoTo Teardown

errores:
    Test_ResponsablePorRol_Calidad_DevuelveUsuariosActivosConRol = BuildFail( _
        "Err " & Err.Number & ": " & Err.Description, m_Logs, "")
Teardown:
    DisposeIsolatedTempDbResponsable db, m_Path
End Function


Public Function Test_ResponsablePorRol_Seguridad_DevuelveUsuariosActivosConRol() As String
    Dim db As DAO.Database
    Dim dic As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Path As String
    Dim m_Logs(0 To 4) As String

    On Error GoTo errores
    m_Path = TEST_DIR & "seguridad_happy_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = CreateIsolatedTempDbResponsable(m_Path)
    SeedRolesFixture db

    m_Logs(0) = "1. Arrange: temp .accdb + seed (3 Seguridad activos)"
    Set dic = modResponsablePorRolHelper.GetUsuariosPorRol( _
                    "Seguridad", "", db, m_Error)
    If m_Error <> "" Then
        Test_ResponsablePorRol_Seguridad_DevuelveUsuariosActivosConRol = BuildFail("unexpected p_Error", m_Logs, m_Error)
        GoTo Teardown
    End If

    m_Logs(1) = "2. Assert: 3 entradas (ids 30006..30008)"
    If dic Is Nothing Then
        Test_ResponsablePorRol_Seguridad_DevuelveUsuariosActivosConRol = BuildFail( _
            "Count=Nothing esperaba 3", m_Logs, "")
        GoTo Teardown
    ElseIf dic.Count <> 3 Then
        Test_ResponsablePorRol_Seguridad_DevuelveUsuariosActivosConRol = BuildFail( _
            "Count=" & CStr(dic.Count) & " esperaba 3", m_Logs, "")
        GoTo Teardown
    End If

    Test_ResponsablePorRol_Seguridad_DevuelveUsuariosActivosConRol = BuildOk("ok", m_Logs)
    GoTo Teardown

errores:
    Test_ResponsablePorRol_Seguridad_DevuelveUsuariosActivosConRol = BuildFail( _
        "Err " & Err.Number & ": " & Err.Description, m_Logs, "")
Teardown:
    DisposeIsolatedTempDbResponsable db, m_Path
End Function


Public Function Test_ResponsablePorRol_Calidad_IncluyeIDGuardadoAunqueSinRol() As String
    Dim db As DAO.Database
    Dim dic As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Path As String
    Dim m_Logs(0 To 4) As String
    Dim m_IDOrfano As Long

    On Error GoTo errores
    m_Path = TEST_DIR & "compat_orfano_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = CreateIsolatedTempDbResponsable(m_Path)
    SeedRolesFixture db

    ' ID 30012 existe en TbUsuariosAplicaciones pero NO tiene rol
    ' asignado (escenario historico: expediente guardo a un usuario que
    ' despues dejo el rol).
    InsertUsuario db, 30012, "orfa@example.com", "orf", _
                  "Usuaria Orfa Sin Rol", "2000-01-01", True, Null
    m_IDOrfano = 30012

    m_Logs(0) = "1. Arrange: temp + seed + usuario 30012 sin rol"
    Set dic = modResponsablePorRolHelper.GetUsuariosPorRol( _
                    "Calidad", CStr(m_IDOrfano), db, m_Error)
    If m_Error <> "" Then
        Test_ResponsablePorRol_Calidad_IncluyeIDGuardadoAunqueSinRol = BuildFail( _
            "unexpected p_Error", m_Logs, m_Error)
        GoTo Teardown
    End If

    m_Logs(1) = "2. Assert: 5 Calidad + 1 orfa incluida = 6 entradas"
    If dic Is Nothing Then
        Test_ResponsablePorRol_Calidad_IncluyeIDGuardadoAunqueSinRol = BuildFail( _
            "Count=Nothing esperaba 6", m_Logs, "")
        GoTo Teardown
    ElseIf dic.Count <> 6 Then
        Test_ResponsablePorRol_Calidad_IncluyeIDGuardadoAunqueSinRol = BuildFail( _
            "Count=" & CStr(dic.Count) & " esperaba 6", m_Logs, "")
        GoTo Teardown
    End If
    If Not dic.Exists(CStr(m_IDOrfano)) Then
        Test_ResponsablePorRol_Calidad_IncluyeIDGuardadoAunqueSinRol = BuildFail( _
            "ID guardado " & m_IDOrfano & " no esta en diccionario", m_Logs, "")
        GoTo Teardown
    End If

    Test_ResponsablePorRol_Calidad_IncluyeIDGuardadoAunqueSinRol = BuildOk("ok", m_Logs)
    GoTo Teardown

errores:
    Test_ResponsablePorRol_Calidad_IncluyeIDGuardadoAunqueSinRol = BuildFail( _
        "Err " & Err.Number & ": " & Err.Description, m_Logs, "")
Teardown:
    DisposeIsolatedTempDbResponsable db, m_Path
End Function


Public Function Test_ResponsablePorRol_RechazaRolFueraDeWhitelist() As String
    Dim db As DAO.Database
    Dim dic As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Path As String
    Dim m_Logs(0 To 4) As String
    Dim m_ErrNum As Long

    On Error GoTo errores
    m_Path = TEST_DIR & "sad_whitelist_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = CreateIsolatedTempDbResponsable(m_Path)
    SeedRolesFixture db

    m_Logs(0) = "1. Arrange: temp + seed"
    On Error Resume Next
    Set dic = modResponsablePorRolHelper.GetUsuariosPorRol( _
                    "Hacker", "", db, m_Error)
    m_ErrNum = Err.Number
    On Error GoTo errores

    m_Logs(1) = "2. Assert: p_Error no vacio Y Err.Raise (1000) propagado"
    If m_Error = "" Then
        Test_ResponsablePorRol_RechazaRolFueraDeWhitelist = BuildFail( _
            "p_Error vacio; whitelist no protegio", m_Logs, "")
        GoTo Teardown
    End If
    If m_ErrNum <> 1000 Then
        Test_ResponsablePorRol_RechazaRolFueraDeWhitelist = BuildFail( _
            "Err=" & m_ErrNum & " esperaba 1000", m_Logs, m_Error)
        GoTo Teardown
    End If

    Test_ResponsablePorRol_RechazaRolFueraDeWhitelist = BuildOk("ok", m_Logs)
    GoTo Teardown

errores:
    Test_ResponsablePorRol_RechazaRolFueraDeWhitelist = BuildFail( _
        "Err " & Err.Number & ": " & Err.Description, m_Logs, "")
Teardown:
    DisposeIsolatedTempDbResponsable db, m_Path
End Function


Public Function Test_ResponsablePorRol_CalidadExcluyeBajaLogica() As String
    Dim db As DAO.Database
    Dim dic As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Path As String
    Dim m_Logs(0 To 4) As String

    On Error GoTo errores
    m_Path = TEST_DIR & "edge_baja_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = CreateIsolatedTempDbResponsable(m_Path)
    SeedRolesFixture db

    ' El usuario 30001 (Ana Calidad) lo "damos de baja" poniendo FechaBaja.
    SetUsuarioBaja db, FIXTURE_ID_BASE + 1

    m_Logs(0) = "1. Arrange: temp + seed + Ana (30001) FechaBaja=today"
    Set dic = modResponsablePorRolHelper.GetUsuariosPorRol( _
                    "Calidad", "", db, m_Error)
    If m_Error <> "" Then
        Test_ResponsablePorRol_CalidadExcluyeBajaLogica = BuildFail( _
            "unexpected p_Error", m_Logs, m_Error)
        GoTo Teardown
    End If

    m_Logs(1) = "2. Assert: 4 entradas (Ana excluida por FechaBaja)"
    If dic Is Nothing Then
        Test_ResponsablePorRol_CalidadExcluyeBajaLogica = BuildFail( _
            "Count=Nothing esperaba 4", m_Logs, "")
        GoTo Teardown
    ElseIf dic.Count <> 4 Then
        Test_ResponsablePorRol_CalidadExcluyeBajaLogica = BuildFail( _
            "Count=" & CStr(dic.Count) & " esperaba 4", m_Logs, "")
        GoTo Teardown
    End If
    If dic.Exists(CStr(FIXTURE_ID_BASE + 1)) Then
        Test_ResponsablePorRol_CalidadExcluyeBajaLogica = BuildFail( _
            "Ana (30001) NO debio aparecer", m_Logs, "")
        GoTo Teardown
    End If

    Test_ResponsablePorRol_CalidadExcluyeBajaLogica = BuildOk("ok", m_Logs)
    GoTo Teardown

errores:
    Test_ResponsablePorRol_CalidadExcluyeBajaLogica = BuildFail( _
        "Err " & Err.Number & ": " & Err.Description, m_Logs, "")
Teardown:
    DisposeIsolatedTempDbResponsable db, m_Path
End Function


Public Function Test_ResponsablePorRol_NoIncluyeUsuarioSinRolEnOtroRol() As String
    ' Cross-rol isolation: un usuario de Seguridad no debe filtrarse cuando
    ' se pide Calidad.  La whitelist por ROL debe ser efectiva.
    Dim db As DAO.Database
    Dim dic As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Path As String
    Dim m_Logs(0 To 4) As String

    On Error GoTo errores
    m_Path = TEST_DIR & "edge_cross_rol_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = CreateIsolatedTempDbResponsable(m_Path)
    SeedRolesFixture db

    m_Logs(0) = "1. Arrange: seed (5 Calidad + 3 Seguridad). Pido Calidad."
    Set dic = modResponsablePorRolHelper.GetUsuariosPorRol( _
                    "Calidad", "", db, m_Error)
    If m_Error <> "" Then
        Test_ResponsablePorRol_NoIncluyeUsuarioSinRolEnOtroRol = BuildFail( _
            "unexpected p_Error", m_Logs, m_Error)
        GoTo Teardown
    End If

    m_Logs(1) = "2. Assert: ninguno de los IDs de Seguridad (30006..30008)"
    If dic Is Nothing Then
        Test_ResponsablePorRol_NoIncluyeUsuarioSinRolEnOtroRol = BuildFail( _
            "dic es Nothing", m_Logs, "")
        GoTo Teardown
    End If
    Dim i As Long
    For i = 6 To 8
        If dic.Exists(CStr(FIXTURE_ID_BASE + i)) Then
            Test_ResponsablePorRol_NoIncluyeUsuarioSinRolEnOtroRol = BuildFail( _
                "ID " & (FIXTURE_ID_BASE + i) & " (Seguridad) entro al pedir Calidad", _
                m_Logs, "")
            GoTo Teardown
        End If
    Next

    Test_ResponsablePorRol_NoIncluyeUsuarioSinRolEnOtroRol = BuildOk("ok", m_Logs)
    GoTo Teardown

errores:
    Test_ResponsablePorRol_NoIncluyeUsuarioSinRolEnOtroRol = BuildFail( _
        "Err " & Err.Number & ": " & Err.Description, m_Logs, "")
Teardown:
    DisposeIsolatedTempDbResponsable db, m_Path
End Function


Public Function Test_ResponsablePorRol_VacioSinUsuariosParaRol_DevuelveDicVacio() As String
    ' Edge: un rol sin usuarios asignados devuelve Dictionary vacio (no
    ' error) para que el combo pueda renderizar solo el sentinel.
    Dim db As DAO.Database
    Dim dic As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Path As String
    Dim m_Logs(0 To 4) As String

    On Error GoTo errores
    m_Path = TEST_DIR & "edge_vacio_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = CreateIsolatedTempDbResponsable(m_Path)
    ' Tablas vacias -- sin seed.
    m_Logs(0) = "1. Arrange: temp con tablas vacias"

    Set dic = modResponsablePorRolHelper.GetUsuariosPorRol( _
                    "Calidad", "", db, m_Error)
    If m_Error <> "" Then
        Test_ResponsablePorRol_VacioSinUsuariosParaRol_DevuelveDicVacio = _
            BuildFail("unexpected p_Error", m_Logs, m_Error)
        GoTo Teardown
    End If
    m_Logs(1) = "2. Assert: dic NO Nothing y Count=0"
    If dic Is Nothing Then
        Test_ResponsablePorRol_VacioSinUsuariosParaRol_DevuelveDicVacio = _
            BuildFail("dic es Nothing (esperaba Dictionary vacio)", m_Logs, "")
        GoTo Teardown
    End If
    If dic.Count <> 0 Then
        Test_ResponsablePorRol_VacioSinUsuariosParaRol_DevuelveDicVacio = _
            BuildFail("Count=" & dic.Count & " esperaba 0", m_Logs, "")
        GoTo Teardown
    End If

    Test_ResponsablePorRol_VacioSinUsuariosParaRol_DevuelveDicVacio = BuildOk("ok", m_Logs)
    GoTo Teardown

errores:
    Test_ResponsablePorRol_VacioSinUsuariosParaRol_DevuelveDicVacio = BuildFail( _
        "Err " & Err.Number & ": " & Err.Description, m_Logs, "")
Teardown:
    DisposeIsolatedTempDbResponsable db, m_Path
End Function


Public Function Test_ResponsablePorRol_UsuarioDuplicadoEnMultiplesRoles_NoSeCuentaDoble() As String
    ' AC7: no duplicates -- el dictionary esta keyed por CStr(Id).  Un
    ' usuario con dos filas de rol en TbResponsablesPorRol solo aparece
    ' una vez en el resultado.
    Dim db As DAO.Database
    Dim dic As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Path As String
    Dim m_Logs(0 To 4) As String

    On Error GoTo errores
    m_Path = TEST_DIR & "edge_duplicado_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = CreateIsolatedTempDbResponsable(m_Path)
    SeedRolesFixture db

    ' Usuario 30011 con doble fila: Calidad + Seguridad.
    InsertUsuario db, FIXTURE_ID_BASE + 11, "dup@example.com", "dup", _
                  "Usuaria Duplicada Cross", "2000-01-01", True, Null
    InsertRol db, FIXTURE_ID_BASE + 11, "Calidad"
    InsertRol db, FIXTURE_ID_BASE + 11, "Seguridad"

    m_Logs(0) = "1. Arrange: temp + seed + usuario 30011 con Calidad Y Seguridad"
    Set dic = modResponsablePorRolHelper.GetUsuariosPorRol( _
                    "Calidad", "", db, m_Error)
    If m_Error <> "" Then
        Test_ResponsablePorRol_UsuarioDuplicadoEnMultiplesRoles_NoSeCuentaDoble = _
            BuildFail("unexpected p_Error", m_Logs, m_Error)
        GoTo Teardown
    End If
    m_Logs(1) = "2. Assert: Count=6 (5 Calidad + 1 cross-rol 30011 sin doble)"
    If dic Is Nothing Then
        Test_ResponsablePorRol_UsuarioDuplicadoEnMultiplesRoles_NoSeCuentaDoble = _
            BuildFail("dic es Nothing", m_Logs, "")
        GoTo Teardown
    End If
    If dic.Count <> 6 Then
        Test_ResponsablePorRol_UsuarioDuplicadoEnMultiplesRoles_NoSeCuentaDoble = _
            BuildFail("Count=" & dic.Count & " esperaba 6", m_Logs, "")
        GoTo Teardown
    End If
    If Not dic.Exists(CStr(FIXTURE_ID_BASE + 11)) Then
        Test_ResponsablePorRol_UsuarioDuplicadoEnMultiplesRoles_NoSeCuentaDoble = _
            BuildFail("30011 ausente del dict", m_Logs, "")
        GoTo Teardown
    End If

    Test_ResponsablePorRol_UsuarioDuplicadoEnMultiplesRoles_NoSeCuentaDoble = _
        BuildOk("ok", m_Logs)
    GoTo Teardown

errores:
    Test_ResponsablePorRol_UsuarioDuplicadoEnMultiplesRoles_NoSeCuentaDoble = _
        BuildFail("Err " & Err.Number & ": " & Err.Description, m_Logs, "")
Teardown:
    DisposeIsolatedTempDbResponsable db, m_Path
End Function


' -----------------------------------------------------------------------------
' Test del service: combina los dos helpers en un envelope JSON.
' -----------------------------------------------------------------------------

Public Function Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos() As String
    Dim db As DAO.Database
    Dim m_Error As String
    Dim m_Path As String
    Dim m_Json As String
    Dim m_Envelope As Object
    Dim m_Payload As Object
    Dim m_ParseErr As Long
    Dim m_Logs(0 To 7) As String

    On Error GoTo errores
    m_Path = TEST_DIR & "service_combobox_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = CreateIsolatedTempDbResponsable(m_Path)
    SeedRolesFixture db

    ' Orphan: 30012 existe en TbUsuariosAplicaciones pero NO tiene rol.
    InsertUsuario db, 30012, "orfa@example.com", "orf", _
                  "Usuaria Orfa Sin Rol", "2000-01-01", True, Null

    ' User con FechaBaja seteada (debe ser excluido de Calidad y Seguridad).
    InsertUsuario db, 30010, "baja@example.com", "baj", _
                  "Usuaria Baja", "2000-01-01", True, CDate("2020-01-01")
    InsertRol db, 30010, "Calidad"

    ' Cross-rol: 30011 esta en Calidad Y Seguridad.
    InsertUsuario db, 30011, "cross@example.com", "cro", _
                  "Usuaria Cross", "2000-01-01", True, Null
    InsertRol db, 30011, "Calidad"
    InsertRol db, 30011, "Seguridad"

    m_Logs(0) = "1. Arrange: temp .accdb (5 Calidad + 3 Seguridad + 1 orphan + 1 baja + 1 cross-rol)"

    m_Json = modResponsablePorRolService.ExpedienteGeneral_EstablecerCombos( _
                Nothing, "30012", "30012", db, m_Error)
    If m_Error <> "" Then
        Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
            BuildFail("service returned p_Error: " & m_Error, m_Logs, "")
        GoTo Teardown
    End If

    m_Logs(1) = "2. Parse envelope via JsonConverter.ParseJson"
    On Error Resume Next
    Set m_Envelope = JsonConverter.ParseJson(m_Json)
    m_ParseErr = Err.Number
    On Error GoTo errores
    If m_Envelope Is Nothing Then
        Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
            BuildFail("ParseJson returned Nothing (Err=" & m_ParseErr & _
                      ") | json=" & Left$(m_Json, 200), m_Logs, "")
        GoTo Teardown
    End If

    m_Logs(2) = "3. Assert: ok=true + payload present"
    If CBool(m_Envelope("ok")) <> True Then
        Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
            BuildFail("ok=" & CStr(m_Envelope("ok")) & _
                      " esperaba true; error=" & CStr(m_Envelope("error")), m_Logs, "")
        GoTo Teardown
    End If
    If Not m_Envelope.Exists("payload") Then
        Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
            BuildFail("payload key missing", m_Logs, "")
        GoTo Teardown
    End If
    Set m_Payload = m_Envelope("payload")

    m_Logs(3) = "4. Assert: calidad.Count = 5 (seed) + 1 (cross 30011) + 1 (orphan 30012) = 7"
    If m_Payload("calidad").Count <> 7 Then
        Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
            BuildFail("calidad.Count=" & m_Payload("calidad").Count & _
                      " esperaba 7", m_Logs, "")
        GoTo Teardown
    End If

    m_Logs(4) = "5. Assert: seguridad.Count = 3 (seed) + 1 (cross 30011) + 1 (orphan 30012) = 5"
    If m_Payload("seguridad").Count <> 5 Then
        Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
            BuildFail("seguridad.Count=" & m_Payload("seguridad").Count & _
                      " esperaba 5", m_Logs, "")
        GoTo Teardown
    End If

    m_Logs(5) = "6. Assert: seguridadSentinel {id:0, label:N/A}"
    If CStr(m_Payload("seguridadSentinel")("id")) <> "0" Or _
       CStr(m_Payload("seguridadSentinel")("label")) <> "N/A" Then
        Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
            BuildFail("sentinel invalido: " & _
                      CStr(m_Payload("seguridadSentinel")("id")) & ";" & _
                      CStr(m_Payload("seguridadSentinel")("label")), m_Logs, "")
        GoTo Teardown
    End If

    m_Logs(6) = "7. Assert: cmbRows.seguridad(1) = '0;N/A' (UX contract primero; Collection VBA es 1-indexed)"
    If CStr(m_Payload("cmbRows")("seguridad")(1)) <> "0;N/A" Then
        Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
            BuildFail("cmbRows.seguridad(1)='" & _
                      CStr(m_Payload("cmbRows")("seguridad")(1)) & _
                      "' esperaba '0;N/A'", m_Logs, "")
        GoTo Teardown
    End If

    m_Logs(7) = "8. Assert: orphanCalidadInList + orphanSeguridadInList = true"
    If CBool(m_Payload("orphanCalidadInList")) <> True Then
        Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
            BuildFail("orphanCalidadInList=false; esperaba true (30012)", m_Logs, "")
        GoTo Teardown
    End If
    If CBool(m_Payload("orphanSeguridadInList")) <> True Then
        Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
            BuildFail("orphanSeguridadInList=false; esperaba true (30012)", m_Logs, "")
        GoTo Teardown
    End If

    Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
        BuildOk("ok", m_Logs)
    GoTo Teardown

errores:
    Test_ResponsablePorRol_Service_ExpedienteGeneral_EstablecerCombos_RetornaAmbosCombos = _
        BuildFail("Err " & Err.Number & ": " & Err.Description, m_Logs, "")
Teardown:
    DisposeIsolatedTempDbResponsable db, m_Path
End Function


' -----------------------------------------------------------------------------
' Helpers privados -- Temp .accdb + seed (sandbox §5.5)
' -----------------------------------------------------------------------------

Private Function CreateIsolatedTempDbResponsable(ByVal p_Path As String) As DAO.Database
    Dim db As DAO.Database
    Dim m_Error As String

    On Error GoTo errores
    If Len(Dir(TEST_DIR, vbDirectory)) = 0 Then MkDir TEST_DIR
    If Len(Dir(p_Path)) > 0 Then Kill p_Path

    Set db = DBEngine.Workspaces(0).CreateDatabase(p_Path, dbLangGeneral, dbVersion120)

    ' TbResponsablesPorRol: COUNTER autonumerico para IDResponsablePorRol (PK)
    ' + IDUsuario Long + Rol Text.  Indice unico (IDUsuario, Rol) como en
    ' produccion.
    db.Execute "CREATE TABLE TbResponsablesPorRol (" & _
                "IDResponsablePorRol COUNTER NOT NULL PRIMARY KEY, " & _
                "IDUsuario INTEGER NOT NULL, " & _
                "Rol TEXT(50) NOT NULL);", dbFailOnError

    db.Execute "CREATE UNIQUE INDEX UX_TbResponsablesPorRol_UsuarioRol " & _
                "ON TbResponsablesPorRol (IDUsuario, Rol);", dbFailOnError

    ' TbUsuariosAplicaciones: Id SHORT (16-bit, tipo real = Integer) + el resto
    ' del subset.  SHORT en DDL de Access corresponde a DAO dbInteger (16 bits
    ' signed, max 32767), NO a INTEGER (que en Access DDL es Long 32 bits).
    db.Execute "CREATE TABLE TbUsuariosAplicaciones (" & _
                "Id SHORT NOT NULL, " & _
                "Nombre TEXT(255), " & _
                "UsuarioRed TEXT(255), " & _
                "CorreoUsuario TEXT(255) NOT NULL, " & _
                "Activado YESNO, " & _
                "FechaBaja DATETIME, " & _
                "FechaAlta DATETIME);", dbFailOnError

    Set CreateIsolatedTempDbResponsable = db
    Exit Function

errores:
    m_Error = "CreateIsolatedTempDbResponsable: " & Err.Number & " - " & Err.Description
    Err.Raise vbObjectError + 513, , m_Error
End Function


Private Sub DisposeIsolatedTempDbResponsable(ByRef p_db As DAO.Database, ByVal p_Path As String)
    On Error Resume Next
    If Not p_db Is Nothing Then p_db.Close
    Set p_db = Nothing
    If Len(Dir(p_Path)) > 0 Then Kill p_Path
    On Error GoTo 0
End Sub


Private Sub SeedRolesFixture(ByRef p_db As DAO.Database)
    ' 5 Calidad activos: 30001..30005
    ' 3 Seguridad activos: 30006..30008
    Dim i As Long
    For i = 1 To 5
        InsertUsuario p_db, FIXTURE_ID_BASE + i, _
                      "calidad" & i & "@example.com", _
                      "cal" & i, _
                      "Usuaria Calidad " & i, _
                      "2000-01-01", _
                      True, _
                      Null
        InsertRol p_db, FIXTURE_ID_BASE + i, "Calidad"
    Next
    For i = 1 To 3
        InsertUsuario p_db, FIXTURE_ID_BASE + 5 + i, _
                      "seguridad" & i & "@example.com", _
                      "seg" & i, _
                      "Usuaria Seguridad " & i, _
                      "2000-01-01", _
                      True, _
                      Null
        InsertRol p_db, FIXTURE_ID_BASE + 5 + i, "Seguridad"
    Next
End Sub


Private Sub InsertUsuario( _
                            ByRef p_db As DAO.Database, _
                            ByVal p_Id As Long, _
                            ByVal p_Correo As String, _
                            ByVal p_UsuarioRed As String, _
                            ByVal p_Nombre As String, _
                            ByVal p_FechaAlta As String, _
                            ByVal p_Activado As Boolean, _
                            ByVal p_FechaBaja As Variant)
    Dim m_SQL As String
    m_SQL = "INSERT INTO TbUsuariosAplicaciones " & _
            "(Id, CorreoUsuario, UsuarioRed, Nombre, FechaAlta, Activado, FechaBaja) " & _
            "VALUES (" & p_Id & ", '" & Replace(p_Correo, "'", "''") & "', " & _
            "'" & Replace(p_UsuarioRed, "'", "''") & "', " & _
            "'" & Replace(p_Nombre, "'", "''") & "', " & _
            "#" & p_FechaAlta & "#, " & BoolSql(p_Activado) & ", " & _
            DateSql(p_FechaBaja) & ");"
    p_db.Execute m_SQL, dbFailOnError
End Sub


Private Sub InsertRol( _
                        ByRef p_db As DAO.Database, _
                        ByVal p_IDUsuario As Long, _
                        ByVal p_Rol As String)
    Dim m_SQL As String
    ' COUNTER autonumerico asigna IDResponsablePorRol automaticamente.
    m_SQL = "INSERT INTO TbResponsablesPorRol " & _
            "(IDUsuario, Rol) VALUES (" & _
            p_IDUsuario & ", '" & _
            Replace(p_Rol, "'", "''") & "');"
    p_db.Execute m_SQL, dbFailOnError
End Sub


Private Sub SetUsuarioBaja(ByRef p_db As DAO.Database, ByVal p_Id As Long)
    p_db.Execute "UPDATE TbUsuariosAplicaciones SET FechaBaja = #" & _
                  Format(Now, "yyyy-mm-dd") & "# WHERE Id = " & p_Id & ";", _
                  dbFailOnError
End Sub


Private Function BoolSql(ByVal p_Valor As Boolean) As String
    If p_Valor Then
        BoolSql = "True"
    Else
        BoolSql = "False"
    End If
End Function


Private Function DateSql(ByVal p_Valor As Variant) As String
    If IsNull(p_Valor) Then
        DateSql = "NULL"
    Else
        DateSql = "#" & Format(CDate(p_Valor), "yyyy-mm-dd") & "#"
    End If
End Function


' -----------------------------------------------------------------------------
' Wrappers JSON (runner contract §2)
' -----------------------------------------------------------------------------

Private Function BuildOk(ByVal p_Value As String, ByRef p_Logs() As String) As String
    BuildOk = "{""ok"":true,""value"":""" & EscapeJson(p_Value) & """,""payload"":null,""error"":null,""logs"":[" & LogsToJson(p_Logs) & "]}"
End Function


Private Function BuildFail(ByVal p_Msg As String, ByRef p_Logs() As String, ByVal p_Det As String) As String
    Dim m_ErrJson As String
    Dim m_MsgCompleto As String
    m_MsgCompleto = p_Msg
    If Len(p_Det) > 0 Then m_MsgCompleto = m_MsgCompleto & " | " & p_Det
    m_ErrJson = """" & EscapeJson(m_MsgCompleto) & """"
    BuildFail = "{""ok"":false,""value"":null,""payload"":null,""error"":" & m_ErrJson & ",""logs"":[" & LogsToJson(p_Logs) & "]}"
End Function


Private Function LogsToJson(ByRef p_Logs() As String) As String
    Dim i As Long, s As String
    For i = LBound(p_Logs) To UBound(p_Logs)
        If Len(p_Logs(i)) > 0 Then
            If Len(s) > 0 Then s = s & ","
            s = s & """" & EscapeJson(p_Logs(i)) & """"
        End If
    Next
    LogsToJson = s
End Function


Private Function EscapeJson(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbCr, "\n")
    EscapeJson = s
End Function