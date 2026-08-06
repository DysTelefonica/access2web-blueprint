Attribute VB_Name = "Test_JerarquiaRolesHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' Test_JerarquiaRolesHelper - Slice 3.3 BR-010 role hierarchy atoms
'
' Each atom is a Public Function (global, unique) that returns the
' canonical JSON via TestHelper.BuildJsonOk / BuildJsonFail. Pure
' helper tests: NO DAO, NO controls, NO MsgBox, no fixtures required.
' ==========================================================================

Public Function Test_JerarquiaRolesHelper_AdministradorPuedeCualquierTransicion() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)

    logs(0) = "1. Arrange: Admin debe poder cualquier (origen -> destino) de los 4 pares representativos"

    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Administrador", "estadoPreregistro", "estadoRegistro") Then _
        Err.Raise 513, , "Admin Preregistro->Registro debe ser True"
    logs(1) = "2. PASS: Admin Preregistro->Registro"

    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Administrador", "estadoModificacion", "estadoValidacion") Then _
        Err.Raise 513, , "Admin Modificacion->Validacion debe ser True"
    logs(2) = "3. PASS: Admin Modificacion->Validacion"

    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Administrador", "estadoValidacion", "estadoRevision") Then _
        Err.Raise 513, , "Admin Validacion->Revision debe ser True"
    logs(3) = "4. PASS: Admin Validacion->Revision"

    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Administrador", "estadoFormalizacion", "estadoRechazada") Then _
        Err.Raise 513, , "Admin Formalizacion->Rechazada debe ser True"
    logs(4) = "5. PASS: Admin Formalizacion->Rechazada"

    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Administrador", "estadoAprobada", "estadoPreregistro") Then _
        Err.Raise 513, , "Admin Aprobada->Preregistro (reapertura) debe ser True"
    logs(5) = "6. PASS: Admin Aprobada->Preregistro"

    Test_JerarquiaRolesHelper_AdministradorPuedeCualquierTransicion = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_JerarquiaRolesHelper_AdministradorPuedeCualquierTransicion = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_JerarquiaRolesHelper_CalidadAValidacionRevision_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)

    logs(0) = "1. Arrange: Calidad transiciones permitidas (BR-010 + spec escenario)"

    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Calidad", "estadoModificacion", "estadoValidacion") Then _
        Err.Raise 513, , "Calidad Modificacion->Validacion debe ser True"
    logs(1) = "2. PASS: Calidad -> estadoValidacion"

    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Calidad", "estadoValidacion", "estadoRevision") Then _
        Err.Raise 513, , "Calidad Validacion->Revision debe ser True"
    logs(2) = "3. PASS: Calidad -> estadoRevision"

    ' Calidad tambien puede devolver la solicitud al tecnico para subsanacion.
    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Calidad", "estadoValidacion", "estadoModificacion") Then _
        Err.Raise 513, , "Calidad puede devolver a estadoModificacion (subsanacion)"
    logs(3) = "4. PASS: Calidad -> estadoModificacion (subsanacion)"

    Test_JerarquiaRolesHelper_CalidadAValidacionRevision_True = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_JerarquiaRolesHelper_CalidadAValidacionRevision_True = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_JerarquiaRolesHelper_TecnicoNoPuedePasarAValidacion_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: Tecnico NO puede transicionar a estadoValidacion (BR-010 + spec)"

    If JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Tecnico", "estadoDesarrolloTecnico", "estadoValidacion") Then _
        Err.Raise 513, , "Tecnico no debe transicionar a estadoValidacion"
    logs(1) = "2. PASS: Tecnico DesarrolloTecnico->Validacion = False"

    ' Variante: tampoco desde estadoModificacion.
    If JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Tecnico", "estadoModificacion", "estadoValidacion") Then _
        Err.Raise 513, , "Tecnico desde estadoModificacion tampoco debe pasar a Validacion"
    logs(2) = "3. PASS: Tecnico Modificacion->Validacion = False"

    Test_JerarquiaRolesHelper_TecnicoNoPuedePasarAValidacion_True = TestHelper.BuildJsonOk("false", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_JerarquiaRolesHelper_TecnicoNoPuedePasarAValidacion_True = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_JerarquiaRolesHelper_TecnicoSubsanaModificacion_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: Tecnico subsana dentro del loop DesarrolloTecnico <-> Modificacion"

    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Tecnico", "estadoModificacion", "estadoDesarrolloTecnico") Then _
        Err.Raise 513, , "Tecnico Modificacion->DesarrolloTecnico debe ser True (subsanacion)"
    logs(1) = "2. PASS: Tecnico Modificacion->DesarrolloTecnico"

    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Tecnico", "estadoDesarrolloTecnico", "estadoModificacion") Then _
        Err.Raise 513, , "Tecnico DesarrolloTecnico->Modificacion debe ser True"
    logs(2) = "3. PASS: Tecnico DesarrolloTecnico->Modificacion"

    Test_JerarquiaRolesHelper_TecnicoSubsanaModificacion_True = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_JerarquiaRolesHelper_TecnicoSubsanaModificacion_True = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_JerarquiaRolesHelper_SinAcceso_Nunca_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)

    logs(0) = "1. Arrange: SinAcceso nunca debe transicionar"

    If JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("SinAcceso", "estadoModificacion", "estadoValidacion") Then _
        Err.Raise 513, , "SinAcceso no debe transicionar Mod->Val"
    logs(1) = "2. PASS: SinAcceso Mod->Val = False"

    If JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("SinAcceso", "estadoDesarrolloTecnico", "estadoModificacion") Then _
        Err.Raise 513, , "SinAcceso no debe transicionar DT->Mod"
    logs(2) = "3. PASS: SinAcceso DT->Mod = False"

    If JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("SinAcceso", "estadoAprobada", "estadoPreregistro") Then _
        Err.Raise 513, , "SinAcceso no debe reabrir"
    logs(3) = "4. PASS: SinAcceso Aprobada->Preregistro = False"

    Test_JerarquiaRolesHelper_SinAcceso_Nunca_True = TestHelper.BuildJsonOk("false", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_JerarquiaRolesHelper_SinAcceso_Nunca_True = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_JerarquiaRolesHelper_RolDesconocido_False() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: roles no catalogados devuelven False (safe-by-default)"

    If JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Patata", "estadoModificacion", "estadoValidacion") Then _
        Err.Raise 513, , "Rol 'Patata' no debe transicionar"
    logs(1) = "2. PASS: 'Patata' -> False"

    If JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("", "estadoModificacion", "estadoValidacion") Then _
        Err.Raise 513, , "Rol vacio debe devolver False"
    logs(2) = "3. PASS: rol vacio -> False"

    Test_JerarquiaRolesHelper_RolDesconocido_False = TestHelper.BuildJsonOk("false", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_JerarquiaRolesHelper_RolDesconocido_False = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_JerarquiaRolesHelper_EstadoDestinoInvalido_False() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(3)

    logs(0) = "1. Arrange: destino no canonico devuelve False aunque el rol sea Admin"

    If JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Administrador", "estadoModificacion", "Patata") Then _
        Err.Raise 513, , "Destino 'Patata' debe devolver False incluso para Admin"
    logs(1) = "2. PASS: Admin -> 'Patata' = False"

    If JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Administrador", "estadoModificacion", "") Then _
        Err.Raise 513, , "Destino vacio debe devolver False"
    logs(2) = "3. PASS: Admin -> '' = False"

    Test_JerarquiaRolesHelper_EstadoDestinoInvalido_False = TestHelper.BuildJsonOk("false", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_JerarquiaRolesHelper_EstadoDestinoInvalido_False = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_JerarquiaRolesHelper_TrimCaseInsensitiveTrue() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)

    logs(0) = "1. Arrange: rol y estado con variaciones de case + espacios"

    ' "admin" (lowercase) reconocido como Admin.
    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("admin", "estadoModificacion", "estadoValidacion") Then _
        Err.Raise 513, , "'admin' lowercase debe ser reconocido como Admin"
    logs(1) = "2. PASS: 'admin' lowercase -> True"

    ' " ADMIN " con espacios.
    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde(" ADMIN ", "estadoModificacion", "estadoValidacion") Then _
        Err.Raise 513, , "' ADMIN ' con espacios debe ser reconocido como Admin"
    logs(2) = "3. PASS: ' ADMIN ' -> True"

    ' "AdMiN" mixto.
    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("AdMiN", "estadoModificacion", "estadoValidacion") Then _
        Err.Raise 513, , "case mixto debe ser reconocido como Admin"
    logs(3) = "4. PASS: 'AdMiN' -> True"

    ' Estado destino en lowercase.
    If Not JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde("Administrador", "estadomodificacion", "estadoVALIDACION") Then _
        Err.Raise 513, , "estados lowercase deben ser aceptados"
    logs(4) = "5. PASS: estados lowercase -> True"

    Test_JerarquiaRolesHelper_TrimCaseInsensitiveTrue = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "ERR " & Err.Number & " - " & Err.Description
    Test_JerarquiaRolesHelper_TrimCaseInsensitiveTrue = TestHelper.BuildJsonFail(Err.Description, logs)
End Function
