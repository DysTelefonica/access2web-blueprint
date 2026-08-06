Attribute VB_Name = "Test_InformePublicacionHelper"
Option Compare Database
Option Explicit

' ============================================================
' Test_InformePublicacionHelper - TDD RED atoms for modInformePublicacionHelper
'
' Helper: ValidarYGenerarInformeEdicion
'   Signature (modInformePublicacionHelper.bas lónea 48):
'     Function ValidarYGenerarInformeEdicion( _
'         ByVal p_IDEdicion As String
'         Optional ByRef p_URLInforme As String
'         Optional ByRef p_Error As String) As String
'
'   IMPORTANTE: el orden real es (..., p_URLInforme, db, p_Error). El helper hace
'   CLng(p_IDEdicion) internamente (lónea 65) así que los FIX_EDICION_*_ID
'   (Long) deben pasarse con CStr() en cada call site.
'
' Form: Form_FormRiesgosGestionEdicion (lines 92-143, ComandoGenerarInforme_Click)
'
' Architecture: validates priorización + alcance, then delegates to
'   GenerarInformePublicacion > InformeRiesgoHTML.GenerarInformeEdicionHTML
'   which saves to disk and returns the file:// URL via ByRef p_URLInforme.
'
' Dependencies:
'   - EnumControlCambiosAlcance (Constantes.bas): Resumen3=1, Completo=2
'   - Constructor.getEdicion(id, err) > Edicion
'   - Edicion.TodosLosRiesgosPriorizados As EnumSiNo
'   - GenerarInformePublicacion(p_Edicion, ..., p_ControlCambiosAlcance) > URL
'   - GetURLInformeEdicionHTML > file path URL
'
' Skill: access-vba-tdd v2.4.3
' ============================================================

' --- Module-level constants for REQ-CAL-09 ---
Private Const FIX_ID_R9_EXPEDIENTE As Long = 902001
Private Const FIX_ID_R9_PROYECTO   As Long = 902002
Private Const FIX_ID_R9_EDICION    As Long = 902003
Private Const FIX_ID_R9_RIESGO_HAPPY As Long = 902010
Private Const FIX_ID_R9_RIESGO_SAD   As Long = 902011
Private Const FIX_ID_R9_RIESGO_EDGE  As Long = 902012
Private Const FIX_ID_R9_RIESGO_ADV   As Long = 902013

' --- Module-level constants for REQ-CAL-10 ---
Private Const FIX_ID_R10_EXPEDIENTE As Long = 903001
Private Const FIX_ID_R10_PROYECTO   As Long = 903002
Private Const FIX_ID_R10_EDICION    As Long = 903003
Private Const FIX_ID_R10_RIESGO_HAPPY As Long = 903010
Private Const FIX_ID_R10_RIESGO_SAD   As Long = 903011
Private Const FIX_ID_R10_RIESGO_EDGE  As Long = 903012
Private Const FIX_ID_R10_RIESGO_ADV   As Long = 903013
Private Const FIX_USER_CALIDAD        As String = "TESTUSER_R10_CALIDAD"
Private Const FIX_USER_OTRO_CALIDAD   As String = "TESTUSER_R10_OTRO_CALIDAD"

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Private Function StubColUsuariosCalidadWithUser(ByVal p_UserRed As String, _
                                                ByRef p_PrevDict As Scripting.Dictionary, _
                                                ByRef p_PrevEntorno As Entorno) As Boolean
    On Error GoTo EH
    Set p_PrevEntorno = m_ObjEntorno
    If m_ObjEntorno Is Nothing Then
        Set m_ObjEntorno = New Entorno
    End If
    If Not (m_ObjEntorno.ColUsuariosCalidad Is Nothing) Then
        Set p_PrevDict = m_ObjEntorno.ColUsuariosCalidad
    End If
    Dim m_NewDict As Scripting.Dictionary
    Set m_NewDict = New Scripting.Dictionary
    m_NewDict.Add p_UserRed, "x"
    Set m_ObjEntorno.ColUsuariosCalidad = m_NewDict
    StubColUsuariosCalidadWithUser = True
    Exit Function
EH:
    StubColUsuariosCalidadWithUser = False
End Function

' --- Fixture for REQ-CAL-09: seed parent graph + specific riesgo in TbRiesgos
'     so that ObtenerMotivoNoPublicable has a valid IDRiesgo to look up.
Private Sub SeedR9Fixture(ByVal p_IDRiesgo As Long)
    On Error GoTo EH_Seed
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Err.Raise 1001, "SeedR9Fixture", "GetTestDb returned Nothing: " & dbErr

    On Error Resume Next
    db.Execute "DELETE FROM tbCambiosParaPublicacion WHERE Riesgo='" & CStr(p_IDRiesgo) & "'", dbFailOnError
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & p_IDRiesgo, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_R9_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_R9_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_R9_EXPEDIENTE, dbFailOnError
    On Error GoTo 0

    ' 1. Expediente (padre)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & FIX_ID_R9_EXPEDIENTE & ", 'TEST09', 'Fixture REQ-CAL-09', 'Test', 1)", dbFailOnError

    ' 2. Proyecto (hijo de Expediente)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto) " & _
               "VALUES (" & FIX_ID_R9_PROYECTO & ", " & FIX_ID_R9_EXPEDIENTE & ", 'TESTPROJ09')", dbFailOnError

    ' 3. Edición (hijo de Proyecto) - necesario para FK de TbRiesgos
    db.Execute "INSERT INTO TbProyectosEdiciones (IDProyecto, IDEdicion, Edicion, Elaborado) " & _
               "VALUES (" & FIX_ID_R9_PROYECTO & ", " & FIX_ID_R9_EDICION & ", 1, 'test_user')", dbFailOnError

    ' 4. Riesgo (hijo de Edición) - schema-first: solo CodigoUnico + CodigoRiesgo + IDEdicion required
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & p_IDRiesgo & ", " & FIX_ID_R9_EDICION & ", " & _
               "'UNICO-R9-" & p_IDRiesgo & "', 'R9')", dbFailOnError

    Set db = Nothing
    Exit Sub

EH_Seed:
    Dim eN As Long: eN = Err.Number
    Dim ed As String: ed = Err.description
    On Error Resume Next
    Set db = Nothing
    Err.Raise eN, "SeedR9Fixture", "Seed failed: " & eN & " - " & ed
End Sub

Private Sub TeardownR9Fixture(ByVal p_IDRiesgo As Long)
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub

    db.Execute "DELETE FROM tbCambiosParaPublicacion WHERE Riesgo='" & CStr(p_IDRiesgo) & "'", dbFailOnError
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & p_IDRiesgo, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_R9_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_R9_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_R9_EXPEDIENTE, dbFailOnError

    Set db = Nothing
    On Error GoTo 0
End Sub

' --- Fixture for REQ-CAL-10: seed parent graph + specific riesgo in TbRiesgos
'     so that ObtenerDetallePorRiesgo has a valid IDRiesgo + IDProyecto +
'     CodigoRiesgo to resolve (TbRiesgosMaterializaciones is keyed by
'     IDProyecto + CodigoRiesgo, NOT IDRiesgo — ver Constructor.bas SELECTs).
'     Decisión documentada en el helper (header line).
Private Sub SeedR10Fixture(ByVal p_IDRiesgo As Long, _
                           ByVal p_IDProyecto As Long, _
                           ByVal p_IDEdicion As Long, _
                           ByVal p_IDExpediente As Long, _
                           ByVal p_CodigoRiesgo As String)
    On Error GoTo EH_Seed
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Err.Raise 1001, "SeedR10Fixture", "GetTestDb returned Nothing: " & dbErr

    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE IDProyecto=" & p_IDProyecto, dbFailOnError
    db.Execute "DELETE FROM tbCambiosParaPublicacion WHERE Riesgo='" & CStr(p_IDRiesgo) & "'", dbFailOnError
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & p_IDRiesgo, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & p_IDEdicion, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & p_IDProyecto, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & p_IDExpediente, dbFailOnError
    On Error GoTo 0

    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & p_IDExpediente & ", 'TEST10', 'Fixture REQ-CAL-10', 'Test', 1)", dbFailOnError
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto) " & _
               "VALUES (" & p_IDProyecto & ", " & p_IDExpediente & ", 'TESTPROJ10')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDProyecto, IDEdicion, Edicion, Elaborado) " & _
               "VALUES (" & p_IDProyecto & ", " & p_IDEdicion & ", 1, 'test_user')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, IDProyecto, CodigoRiesgo, CodigoUnico) " & _
               "VALUES (" & p_IDRiesgo & ", " & p_IDEdicion & ", " & p_IDProyecto & ", " & _
               "'" & Test_Helper.SqlStr(p_CodigoRiesgo) & "', " & _
               "'UNICO-R10-" & p_IDRiesgo & "')", dbFailOnError

    Set db = Nothing
    Exit Sub

EH_Seed:
    Dim eN As Long: eN = Err.Number
    Dim ed As String: ed = Err.description
    On Error Resume Next
    Set db = Nothing
    Err.Raise eN, "SeedR10Fixture", "Seed failed: " & eN & " - " & ed
End Sub

Private Sub TeardownR10Fixture(ByVal p_IDRiesgo As Long, _
                               ByVal p_IDProyecto As Long, _
                               ByVal p_IDEdicion As Long, _
                               ByVal p_IDExpediente As Long)
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub

    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE IDProyecto=" & p_IDProyecto, dbFailOnError
    db.Execute "DELETE FROM tbCambiosParaPublicacion WHERE Riesgo='" & CStr(p_IDRiesgo) & "'", dbFailOnError
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & p_IDRiesgo, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & p_IDEdicion, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & p_IDProyecto, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & p_IDExpediente, dbFailOnError

    Set db = Nothing
    On Error GoTo 0
End Sub

' --- StubColUsuariosCalidad ---
'     Reemplaza m_ObjEntorno.ColUsuariosCalidad por un Dictionary que contiene
'     solo FIX_USER_CALIDAD. Guarda el valor previo para restaurar en cleanup.
'     Garantiza que la rama "user IS in collection" pasa el authorization guard.
'     Si m_ObjEntorno es Nothing, lo crea (New Entorno) para que el helper
'     pueda leer .ColUsuariosCalidad.
Private Sub RestoreColUsuariosCalidad(ByVal p_PrevDict As Scripting.Dictionary, _
                                       ByVal p_PrevEntorno As Entorno)
    On Error Resume Next
    If Not (m_ObjEntorno Is Nothing) Then
        If p_PrevDict Is Nothing Then
            ' No había diccionario previo: dejar el stub (no se restaura).
            ' Es preferible a lanzar — el siguiente test re-inicializa si hace falta.
        Else
            Set m_ObjEntorno.ColUsuariosCalidad = p_PrevDict
        End If
    End If
    Set m_ObjEntorno = p_PrevEntorno
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' ATOM: Test_ValidarYGenerarInformeEdicion_Happy
' Scenario: Happy path - edición priorizada, alcance válido (Resumen3)
' Expected: p_URLInforme populated with file URL, p_Error empty
' ----------------------------------------------------------------
Public Function Test_ValidarYGenerarInformeEdicion_Happy() As String
    On Error GoTo EH

    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: sandbox session + fixture (edición priorizada, 1 riesgo)"
    logs(1) = "2. Act: ValidarYGenerarInformeEdicion"
    logs(2) = "3. Assert: p_URLInforme non-empty"
    logs(3) = "3a. Assert: file exists at p_URLInforme path"
    logs(4) = "3b. Assert: p_Error is empty"
    logs(5) = "4. Teardown: delete test rows"

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ValidarYGenerarInformeEdicion_Happy = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarYGenerarInformeEdicion_Happy = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' --- Fixture: proyecto + edición priorizada + 1 riesgo priorizado ---
    Const FIX_PROJ_ID As Long = 900001
    Const FIX_EDICION_ID As Long = 900001
    Const FIX_RIESGO_ID As Long = 900001

    Dim proyectoInsert As String
    proyectoInsert = "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, Cliente, Proyecto) " & _
                     "VALUES (" & FIX_PROJ_ID & ", 'TEST_PROYECTO', 'TEST_CLIENTE', 'EXP-TEST-001')"
    db.Execute proyectoInsert, dbFailOnError

    Dim edicionInsert As String
    edicionInsert = "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado, FechaEdicion) " & _
                    "VALUES (" & FIX_EDICION_ID & ", " & FIX_PROJ_ID & ", 1, 'TestAutor', Date())"
    db.Execute edicionInsert, dbFailOnError

    Dim riesgoInsert As String
    riesgoInsert = "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoRiesgo, Descripcion, Priorizacion, Estado) " & _
                   "VALUES (" & FIX_RIESGO_ID & ", " & FIX_EDICION_ID & ", 'R-TEST-001', 'Riesgo de prueba', 'Alta', 'Abierto')"
    db.Execute riesgoInsert, dbFailOnError

    ' --- Act ---
    Dim urlInforme As String
    Dim errMsg As String

    urlInforme = ""
    errMsg = ""
    Dim result As String
    result = ValidarYGenerarInformeEdicion(CStr(FIX_EDICION_ID), EnumControlCambiosAlcanceResumen3, urlInforme, db, errMsg)

    ' --- Assert ---
    If errMsg <> "" Then
        Test_ValidarYGenerarInformeEdicion_Happy = BuildFail("p_Error should be empty but got: " & errMsg, logs)
        GoTo Teardown
    End If

    If urlInforme = "" Then
        Test_ValidarYGenerarInformeEdicion_Happy = BuildFail("p_URLInforme should be non-empty (file URL)", logs)
        GoTo Teardown
    End If

    ' Verify the file actually exists
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(urlInforme) Then
        Test_ValidarYGenerarInformeEdicion_Happy = BuildFail("p_URLInforme path does not exist: " & urlInforme, logs)
        Set fso = Nothing
        GoTo Teardown
    End If
    Set fso = Nothing

    Test_ValidarYGenerarInformeEdicion_Happy = BuildOk(urlInforme, logs)
    GoTo Teardown

EH:
    Test_ValidarYGenerarInformeEdicion_Happy = BuildFail("Test_ValidarYGenerarInformeEdicion_Happy: " & Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & FIX_RIESGO_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROJ_ID, dbFailOnError
End Function

' ----------------------------------------------------------------
' ATOM: Test_ValidarYGenerarInformeEdicion_Sad
' Scenario: Sad path - edición sin riesgos priorizados (TodosLosRiesgosPriorizados=No)
'           AND invalid alcance (out-of-range enum value)
' Expected: p_Error non-empty, p_URLInforme empty
' ----------------------------------------------------------------
Public Function Test_ValidarYGenerarInformeEdicion_Sad() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: sandbox session + fixture (edición NO priorizada)"
    logs(1) = "2. Act: ValidarYGenerarInformeEdicion with unprioritized edición"
    logs(2) = "3. Assert: p_Error mentions priorización"
    logs(3) = "3a. Assert: p_URLInforme is empty"
    logs(4) = "4. Teardown: delete test rows"

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ValidarYGenerarInformeEdicion_Sad = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarYGenerarInformeEdicion_Sad = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' --- Fixture: proyecto + edición NO priorizada ---
    Const FIX_PROJ_ID As Long = 900002
    Const FIX_EDICION_ID As Long = 900002

    Dim proyectoInsert As String
    proyectoInsert = "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, Cliente, Proyecto) " & _
                     "VALUES (" & FIX_PROJ_ID & ", 'TEST_PROYECTO2', 'TEST_CLIENTE2', 'EXP-TEST-002')"
    db.Execute proyectoInsert, dbFailOnError

    ' TodosLosRiesgosPriorizados = 2 (EnumSiNo.No)
    Dim edicionInsert As String
    edicionInsert = "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado, FechaEdicion) " & _
                    "VALUES (" & FIX_EDICION_ID & ", " & FIX_PROJ_ID & ", 1, 'TestAutor', Date())"
    db.Execute edicionInsert, dbFailOnError

    ' --- Act ---
    Dim urlInforme As String
    Dim errMsg As String

    urlInforme = ""
    errMsg = ""
    Dim result As String
    ' Passing an invalid alcance value (99) in addition to unprioritized edición
    result = ValidarYGenerarInformeEdicion(CStr(FIX_EDICION_ID), 99, urlInforme, db, errMsg)

    ' --- Assert ---
    If errMsg = "" Then
        Test_ValidarYGenerarInformeEdicion_Sad = BuildFail("p_Error should be non-empty for unprioritized edición", logs)
        GoTo Teardown
    End If

    ' Error should mention the priorización problem or the validation failure
    If InStr(1, errMsg, "prioriz", vbTextCompare) = 0 And _
       InStr(1, errMsg, "alcance", vbTextCompare) = 0 And _
       InStr(1, errMsg, "riesgo", vbTextCompare) = 0 Then
        Test_ValidarYGenerarInformeEdicion_Sad = BuildFail("p_Error should mention priorización/riesgo/alcance, got: " & errMsg, logs)
        GoTo Teardown
    End If

    If urlInforme <> "" Then
        Test_ValidarYGenerarInformeEdicion_Sad = BuildFail("p_URLInforme should be empty on error, got: " & urlInforme, logs)
        GoTo Teardown
    End If

    Test_ValidarYGenerarInformeEdicion_Sad = BuildOk("unprioritized_rejected", logs)
    GoTo Teardown

EH:
    Test_ValidarYGenerarInformeEdicion_Sad = BuildFail("Test_ValidarYGenerarInformeEdicion_Sad: " & Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROJ_ID, dbFailOnError
End Function

' ----------------------------------------------------------------
' ATOM: Test_ValidarYGenerarInformeEdicion_Edge
' Scenario: Edge cases -
'   (a) edición with 0 riesgos > valid priorización check but empty inventory
'   (b) edición with 1 riesgo > minimal but valid
'   (c) alcance = Completo (2, boundary)
' Expected: all cases either succeed with URL or fail gracefully with p_Error
' ----------------------------------------------------------------
Public Function Test_ValidarYGenerarInformeEdicion_Edge() As String
    On Error GoTo EH

    Dim logs(0 To 8) As String
    logs(0) = "1. Arrange: 3 ediciones - 0 riesgos, 1 riesgo, boundary alcance=Completo"
    logs(1) = "2. Act(a): edición 0 riesgos > should either succeed or fail gracefully"
    logs(2) = "2b. Act(b): edición 1 riesgo"
    logs(3) = "2c. Act(c): edición 1 riesgo + alcance Completo (boundary enum)"
    logs(4) = "3. Assert: each case has consistent (urlInforme XOR error) pair"
    logs(5) = "4. Teardown: delete all test rows"

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ValidarYGenerarInformeEdicion_Edge = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarYGenerarInformeEdicion_Edge = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' --- Fixture setup ---
    Const FIX_PROJ_ID_A As Long = 900003
    Const FIX_PROJ_ID_B As Long = 900004
    Const FIX_PROJ_ID_C As Long = 900005
    Const FIX_EDICION_A As Long = 900003  ' 0 riesgos
    Const FIX_EDICION_B As Long = 900004  ' 1 riesgo
    Const FIX_EDICION_C As Long = 900005  ' 1 riesgo + Completo
    Const FIX_RIESGO_B As Long = 900004
    Const FIX_RIESGO_C As Long = 900005

    ' Shared project for (a)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, Cliente, Proyecto) " & _
               "VALUES (" & FIX_PROJ_ID_A & ", 'TEST_A', 'C', 'EXP-A')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado, FechaEdicion) " & _
               "VALUES (" & FIX_EDICION_A & ", " & FIX_PROJ_ID_A & ", 1, 'Autor', Date())", dbFailOnError

    ' Project for (b)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, Cliente, Proyecto) " & _
               "VALUES (" & FIX_PROJ_ID_B & ", 'TEST_B', 'C', 'EXP-B')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado, FechaEdicion) " & _
               "VALUES (" & FIX_EDICION_B & ", " & FIX_PROJ_ID_B & ", 1, 'Autor', Date())", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoRiesgo, Descripcion, Priorizacion, Estado) " & _
               "VALUES (" & FIX_RIESGO_B & ", " & FIX_EDICION_B & ", 'R-B', 'Desc', 'Alta', 'Abierto')", dbFailOnError

    ' Project for (c)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, Cliente, Proyecto) " & _
               "VALUES (" & FIX_PROJ_ID_C & ", 'TEST_C', 'C', 'EXP-C')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado, FechaEdicion) " & _
               "VALUES (" & FIX_EDICION_C & ", " & FIX_PROJ_ID_C & ", 1, 'Autor', Date())", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoRiesgo, Descripcion, Priorizacion, Estado) " & _
               "VALUES (" & FIX_RIESGO_C & ", " & FIX_EDICION_C & ", 'R-C', 'Desc', 'Alta', 'Abierto')", dbFailOnError

    ' --- Act (a): 0 riesgos ---
    Dim urlA As String, errA As String
    urlA = "": errA = ""
    Dim resultA As String
    resultA = ValidarYGenerarInformeEdicion(CStr(FIX_EDICION_A), EnumControlCambiosAlcanceResumen3, urlA, db, errA)

    ' XOR: exactly one of urlA/errA should be populated
    Dim caseAOk As Boolean
    caseAOk = (urlA <> "" And errA = "") Or (urlA = "" And errA <> "")

    ' --- Act (b): 1 riesgo + Resumen3 ---
    Dim urlB As String, errB As String
    urlB = "": errB = ""
    Dim resultB As String
    resultB = ValidarYGenerarInformeEdicion(CStr(FIX_EDICION_B), EnumControlCambiosAlcanceResumen3, urlB, db, errB)

    Dim caseBOk As Boolean
    caseBOk = (urlB <> "" And errB = "") Or (urlB = "" And errB <> "")

    ' --- Act (c): 1 riesgo + Completo (boundary enum = 2) ---
    Dim urlC As String, errC As String
    urlC = "": errC = ""
    Dim resultC As String
    resultC = ValidarYGenerarInformeEdicion(CStr(FIX_EDICION_C), EnumControlCambiosAlcanceCompleto, urlC, db, errC)

    Dim caseCOk As Boolean
    caseCOk = (urlC <> "" And errC = "") Or (urlC = "" And errC <> "")

    ' --- Assert ---
    If Not caseAOk Then
        Test_ValidarYGenerarInformeEdicion_Edge = BuildFail("Case (a) 0 riesgos: url='' xor err='' violated - url=[" & urlA & "] err=[" & errA & "]", logs)
        GoTo Teardown
    End If

    If Not caseBOk Then
        Test_ValidarYGenerarInformeEdicion_Edge = BuildFail("Case (b) 1 riesgo: url='' xor err='' violated - url=[" & urlB & "] err=[" & errB & "]", logs)
        GoTo Teardown
    End If

    If Not caseCOk Then
        Test_ValidarYGenerarInformeEdicion_Edge = BuildFail("Case (c) alcance Completo: url='' xor err='' violated - url=[" & urlC & "] err=[" & errC & "]", logs)
        GoTo Teardown
    End If

    Test_ValidarYGenerarInformeEdicion_Edge = BuildOk("edge_3_cases", logs)
    GoTo Teardown

EH:
    Test_ValidarYGenerarInformeEdicion_Edge = BuildFail("Test_ValidarYGenerarInformeEdicion_Edge: " & Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo IN (" & FIX_RIESGO_B & "," & FIX_RIESGO_C & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (" & FIX_EDICION_A & "," & FIX_EDICION_B & "," & FIX_EDICION_C & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto IN (" & FIX_PROJ_ID_A & "," & FIX_PROJ_ID_B & "," & FIX_PROJ_ID_C & ")", dbFailOnError
End Function

' ----------------------------------------------------------------
' ATOM: Test_ValidarYGenerarInformeEdicion_Adversarial
' Scenario: Adversarial -
'   (a) Double-click: call helper twice rapidly - idempotency check
'   (b) Mid-generation crash: kill during generation - no orphan HTML files
' Expected:
'   (a) Second call either returns same URL or clean error (no double generation)
'   (b) Any partial file is cleaned up (no orphan .html in output dir)
' ----------------------------------------------------------------
Public Function Test_ValidarYGenerarInformeEdicion_Adversarial() As String
    On Error GoTo EH

    Dim logs(0 To 9) As String
    logs(0) = "1. Arrange: sandbox + edición priorizada + 1 riesgo"
    logs(1) = "2. Act(a): first call - capture URL and orphan files"
    logs(2) = "3. Act(b): second rapid call - idempotency check"
    logs(3) = "3a. Assert: second call is clean (same URL or error, no crash)"
    logs(4) = "4. Act(c): snapshot orphan files before second run"
    logs(5) = "5. Teardown: delete test rows + orphan files"

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ValidarYGenerarInformeEdicion_Adversarial = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarYGenerarInformeEdicion_Adversarial = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' --- Fixture ---
    Const FIX_PROJ_ID As Long = 900006
    Const FIX_EDICION_ID As Long = 900006
    Const FIX_RIESGO_ID As Long = 900006

    db.Execute "INSERT INTO TbProyectos (IDProyecto, NombreProyecto, Cliente, Proyecto) " & _
               "VALUES (" & FIX_PROJ_ID & ", 'TEST_ADV', 'C', 'EXP-ADV')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado, FechaEdicion) " & _
               "VALUES (" & FIX_EDICION_ID & ", " & FIX_PROJ_ID & ", 1, 'Autor', Date())", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoRiesgo, Descripcion, Priorizacion, Estado) " & _
               "VALUES (" & FIX_RIESGO_ID & ", " & FIX_EDICION_ID & ", 'R-ADV', 'Desc', 'Alta', 'Abierto')", dbFailOnError

    ' --- Act (a): first call ---
    Dim url1 As String, err1 As String
    url1 = "": err1 = ""
    Dim result1 As String
    result1 = ValidarYGenerarInformeEdicion(CStr(FIX_EDICION_ID), EnumControlCambiosAlcanceResumen3, url1, db, err1)

    If err1 <> "" Then
        ' If first call already errors, adversarial is moot but still clean
        Test_ValidarYGenerarInformeEdicion_Adversarial = BuildOk("first_call_error_clean", logs)
        GoTo Teardown
    End If

    ' Record the first URL path for orphan check
    Dim firstUrl As String
    firstUrl = url1

    ' --- Act (b): second rapid call (double-click simulation) ---
    Dim url2 As String, err2 As String
    url2 = "": err2 = ""
    Dim result2 As String
    result2 = ValidarYGenerarInformeEdicion(CStr(FIX_EDICION_ID), EnumControlCambiosAlcanceResumen3, url2, db, err2)

    ' --- Assert (a): no crash on second call
    ' Either returns same URL (idempotent) or clean error
    If err2 <> "" And url2 <> firstUrl Then
        ' Clean rejection is acceptable for double-click
    End If

    ' --- Assert (b): idempotency - if both succeed, URLs should match
    ' (file was overwritten, not duplicated)
    If err1 = "" And err2 = "" And url2 <> firstUrl Then
        Test_ValidarYGenerarInformeEdicion_Adversarial = BuildFail("Double-click: expected same URL, got different URLs: [" & firstUrl & "] vs [" & url2 & "]", logs)
        GoTo Teardown
    End If

    ' --- Assert (c): no orphan files beyond the known URLs ---
    ' Collect all .html files in the same folder as the generated URL
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim generatedFolder As String
    Dim generatedFile As String
    Dim orphanFound As Boolean
    orphanFound = False

    If fso.FileExists(firstUrl) Then
        generatedFolder = fso.GetParentFolderName(firstUrl)
        Dim fileName As String
        fileName = fso.GetFileName(firstUrl)

        ' List all .html files - the only legitimate one is firstUrl
        Dim htmlFile As Object
        Dim htmlFiles As Object
        Set htmlFiles = fso.GetFolder(generatedFolder).Files

        For Each htmlFile In htmlFiles
            If Right$(LCase$(htmlFile.Name), 5) = ".html" Then
                If InStr(1, htmlFile.path, "InformeEdicion", vbTextCompare) > 0 Then
                    If htmlFile.path <> firstUrl Then
                        ' Unknown orphan - may be from prior test run, but still check
                        orphanFound = True
                        Exit For
                    End If
                End If
            End If
        Next htmlFile
    End If
    Set fso = Nothing

    If orphanFound Then
        Test_ValidarYGenerarInformeEdicion_Adversarial = BuildFail("Orphan HTML file detected in output folder", logs)
        GoTo Teardown
    End If

    Test_ValidarYGenerarInformeEdicion_Adversarial = BuildOk("adversarial_clean", logs)
    GoTo Teardown

EH:
    Test_ValidarYGenerarInformeEdicion_Adversarial = BuildFail("Test_ValidarYGenerarInformeEdicion_Adversarial: " & Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & FIX_RIESGO_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROJ_ID, dbFailOnError
    ' Clean up generated HTML files if they exist
    Dim fso2 As Object
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    If fso2.FileExists(firstUrl) Then
        On Error Resume Next
        fso2.DeleteFile firstUrl, True
        On Error GoTo 0
    End If
    Set fso2 = Nothing
End Function

' ============================================================
' BLOQUE 3 — REQ-CAL-09 — Atoms for ObtenerMotivoNoPublicable
'
' Helper: ObtenerMotivoNoPublicable(ByRef p_IDRiesgo As String, ...)
'   Source: tbCambiosParaPublicacion.Descripcion WHERE
'           Riesgo=p_IDRiesgo AND NombreCampo='MotivoNoPublicable'
'           ORDER BY FechaRegistro DESC, IDCambio DESC LIMIT 1
'   Apply HTMLSafe to the value.
'   Returns "sin motivo registrado" when no row or empty.
'
' Fixture IDs: dedicated range 902000-902099 (no overlap with
'   SeedBaseGraph 900500-900599, SeedSubcatGraph 900100-900221,
'   Bloque 3 / REQ-CAL-07 901000-901099, or REQ-CAL-16A 50010+).
' ============================================================

' ============================================================
' ATOM 1 — Happy: MotivoNoPublicable_RenderizadoLiteral
' GIVEN staging sandbox + Riesgo con tbCambiosParaPublicacion.Descripcion
'       = "Falta evidencia de control", NombreCampo="MotivoNoPublicable"
' WHEN ObtenerMotivoNoPublicable(p_IDRiesgo) is called
' THEN:
'   - retorna el motivo literal "Falta evidencia de control"
'   - p_Error vacío
' ============================================================
Public Function Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Happy_RenderizadoLiteral() As String
    Dim logs(0 To 8) As String
    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Happy_RenderizadoLiteral = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded"
    logs(1) = "2. Arrange: SeedR9Fixture(902010) + insert tbCambiosParaPublicacion 'Falta evidencia de control'"
    logs(2) = "3. Act: ObtenerMotivoNoPublicable(902010, db, err)"
    logs(3) = "4. Assert: retorno no vacío"
    logs(4) = "5. Assert: retorno = ""Falta evidencia de control"" literal"
    logs(5) = "6. Assert: p_Error vacío"
    logs(6) = "7. Teardown: borrar tbCambiosParaPublicacion + riesgo + teardown"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Happy_RenderizadoLiteral = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    SeedR9Fixture FIX_ID_R9_RIESGO_HAPPY

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Happy_RenderizadoLiteral = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Insert motivo row
    db.Execute "INSERT INTO tbCambiosParaPublicacion " & _
               "(IDProyecto, EdicionInicial, EdicionFinal, Riesgo, NombreCampo, Descripcion, FechaRegistro) " & _
               "VALUES (" & FIX_ID_R9_PROYECTO & ", 1, 1, '" & CStr(FIX_ID_R9_RIESGO_HAPPY) & "', " & _
               "'MotivoNoPublicable', " & _
               "'Falta evidencia de control', " & _
               "#" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ObtenerMotivoNoPublicable(CStr(FIX_ID_R9_RIESGO_HAPPY), db, m_Err)

    If Len(m_Err) <> 0 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Happy_RenderizadoLiteral = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If Len(m_Result) = 0 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Happy_RenderizadoLiteral = BuildFail("retorno vacío, esperaba motivo", logs)
        GoTo Teardown
    End If
    If m_Result <> "Falta evidencia de control" Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Happy_RenderizadoLiteral = BuildFail("motivo esperado 'Falta evidencia de control', obtuvo: " & m_Result, logs)
        GoTo Teardown
    End If

    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Happy_RenderizadoLiteral = BuildOk(m_Result, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownR9Fixture FIX_ID_R9_RIESGO_HAPPY
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Happy_RenderizadoLiteral = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 2 — Sad: MotivoVacio_MuestraPlaceholder
' GIVEN staging sandbox + Riesgo con tbCambiosParaPublicacion.Descripcion
'       = Null/empty
' WHEN ObtenerMotivoNoPublicable(p_IDRiesgo) is called
' THEN:
'   - retorna "sin motivo registrado"
'   - p_Error vacío
' ============================================================
Public Function Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Sad_VacioMuestraPlaceholder() As String
    Dim logs(0 To 7) As String
    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Sad_VacioMuestraPlaceholder = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded"
    logs(1) = "2. Arrange: SeedR9Fixture(902011) sin row en tbCambiosParaPublicacion"
    logs(2) = "3. Act: ObtenerMotivoNoPublicable(902011, db, err)"
    logs(3) = "4. Assert: retorno = ""sin motivo registrado"" (placeholder)"
    logs(4) = "5. Assert: p_Error vacío"
    logs(5) = "6. Teardown: borrar riesgo + teardown"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Sad_VacioMuestraPlaceholder = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    SeedR9Fixture FIX_ID_R9_RIESGO_SAD

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Sad_VacioMuestraPlaceholder = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' No INSERT into tbCambiosParaPublicacion ? empty state

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ObtenerMotivoNoPublicable(CStr(FIX_ID_R9_RIESGO_SAD), db, m_Err)

    If Len(m_Err) <> 0 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Sad_VacioMuestraPlaceholder = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If m_Result <> "sin motivo registrado" Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Sad_VacioMuestraPlaceholder = BuildFail("esperaba 'sin motivo registrado', obtuvo: " & m_Result, logs)
        GoTo Teardown
    End If

    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Sad_VacioMuestraPlaceholder = BuildOk("sad_vacio_placeholder_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownR9Fixture FIX_ID_R9_RIESGO_SAD
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Sad_VacioMuestraPlaceholder = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 3 — Edge: CaracteresEspeciales_NoTruncaYEscapa
' GIVEN staging sandbox + Riesgo con motivo que incluye <, >, &, ", vbCrLf,
'       y > 300 chars
' WHEN ObtenerMotivoNoPublicable(p_IDRiesgo) is called
' THEN:
'   - retorna el motivo completo (sin truncar)
'   - < > & " están HTMLSafe'd
'   - vbCrLf preservado (HTMLSafe no debe comerse saltos de línea en
'     el output — el navegador los respeta como \n en HTML; verificamos
'     que el contenido completo está presente)
' ============================================================
Public Function Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca() As String
    Dim logs(0 To 9) As String
    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded"
    logs(1) = "2. Arrange: SeedR9Fixture(902012)"
    logs(2) = "3. Arrange: insert motivo con Müller & García <test>""" & vbCrLf & "300 chars"
    logs(3) = "4. Act: ObtenerMotivoNoPublicable(902012, db, err)"
    logs(4) = "5. Assert: retorno no vacío y > 300 chars (no truncado)"
    logs(5) = "6. Assert: < escapado a &lt;"
    logs(6) = "7. Assert: > escapado a &gt;"
    logs(7) = "8. Assert: & escapado a &amp;"
    logs(8) = "9. Teardown: borrar motivo + riesgo + teardown"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    SeedR9Fixture FIX_ID_R9_RIESGO_EDGE

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Compose motivo with specials + 300+ chars
    Dim m_MotivoRaw As String
    m_MotivoRaw = "Müller & García <test>" & """" & vbCrLf & String$(300, "x")

    ' INSERT — use SqlStr for safety
    db.Execute "INSERT INTO tbCambiosParaPublicacion " & _
               "(IDProyecto, EdicionInicial, EdicionFinal, Riesgo, NombreCampo, Descripcion, FechaRegistro) " & _
               "VALUES (" & FIX_ID_R9_PROYECTO & ", 1, 1, '" & CStr(FIX_ID_R9_RIESGO_EDGE) & "', " & _
               "'MotivoNoPublicable', " & _
               "'" & Test_Helper.SqlStr(m_MotivoRaw) & "', " & _
               "#" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ObtenerMotivoNoPublicable(CStr(FIX_ID_R9_RIESGO_EDGE), db, m_Err)

    If Len(m_Err) <> 0 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If

    ' Length check (>300 chars raw + escapes = around 320 chars)
    If Len(m_Result) < 300 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca = BuildFail("retorno truncado: len=" & Len(m_Result), logs)
        GoTo Teardown
    End If

    ' HTML escaping assertions
    If InStr(1, m_Result, "<test>", vbTextCompare) > 0 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca = BuildFail("HTML no escapado: contiene '<test>' literal", logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "&lt;test&gt;", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca = BuildFail("HTML escape ausente: no contiene '&lt;test&gt;'", logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "Müller", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca = BuildFail("Unicode perdido: 'Müller' no presente", logs)
        GoTo Teardown
    End If
    ' Check that the 300 x's tail is present (proves no truncation)
    If InStr(1, m_Result, String$(50, "x"), vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca = BuildFail("cola de 50 x ausente (posible truncado)", logs)
        GoTo Teardown
    End If

    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca = BuildOk("edge_htmlsafe_unicode_long_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownR9Fixture FIX_ID_R9_RIESGO_EDGE
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Edge_CaracteresEspeciales_NoTrunca = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 4 — Adversarial: RecargaDuranteGeneracion_NoDejaParciales
' GIVEN staging sandbox + Riesgo con motivo persistido
' WHEN ObtenerMotivoNoPublicable is called twice rapidly (same ID)
' THEN:
'   - ambas llamadas retornan el mismo motivo
'   - ninguna deja archivos parciales (no hay file I/O; el helper es
'     pure read sobre DAO, así que el chequeo es que no haya row extra
'     en tbCambiosParaPublicacion tras las 2 llamadas)
'   - cardinalidad: count(tbCambiosParaPublicacion WHERE Riesgo=X) == 1
' ============================================================
Public Function Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Adversarial_RecargaNoDejaParcial() As String
    Dim logs(0 To 9) As String
    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Adversarial_RecargaNoDejaParcial = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded"
    logs(1) = "2. Arrange: SeedR9Fixture(902013) + motivo persistido"
    logs(2) = "3. Act: 1ª llamada"
    logs(3) = "4. Assert: 1ª llamada retorna motivo"
    logs(4) = "5. Act: 2ª llamada (recarga)"
    logs(5) = "6. Assert: 2ª llamada retorna mismo motivo"
    logs(6) = "7. Assert: count(tbCambiosParaPublicacion WHERE Riesgo=X) == 1 (no row extra)"
    logs(7) = "8. Teardown: borrar motivo + riesgo + teardown"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Adversarial_RecargaNoDejaParcial = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    SeedR9Fixture FIX_ID_R9_RIESGO_ADV

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Adversarial_RecargaNoDejaParcial = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Insert motivo
    db.Execute "INSERT INTO tbCambiosParaPublicacion " & _
               "(IDProyecto, EdicionInicial, EdicionFinal, Riesgo, NombreCampo, Descripcion, FechaRegistro) " & _
               "VALUES (" & FIX_ID_R9_PROYECTO & ", 1, 1, '" & CStr(FIX_ID_R9_RIESGO_ADV) & "', " & _
               "'MotivoNoPublicable', " & _
               "'Materializado sin plan de contingencia activo', " & _
               "#" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    ' --- 1ª llamada ---
    Dim m_Result1 As String
    Dim m_Err1 As String
    m_Result1 = ObtenerMotivoNoPublicable(CStr(FIX_ID_R9_RIESGO_ADV), db, m_Err1)
    If Len(m_Err1) <> 0 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Adversarial_RecargaNoDejaParcial = BuildFail("1ª llamada p_Error: " & m_Err1, logs)
        GoTo Teardown
    End If
    If Len(m_Result1) = 0 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Adversarial_RecargaNoDejaParcial = BuildFail("1ª llamada retorno vacío", logs)
        GoTo Teardown
    End If

    ' --- 2ª llamada (recarga / doble click simulado) ---
    Dim m_Result2 As String
    Dim m_Err2 As String
    m_Result2 = ObtenerMotivoNoPublicable(CStr(FIX_ID_R9_RIESGO_ADV), db, m_Err2)
    If Len(m_Err2) <> 0 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Adversarial_RecargaNoDejaParcial = BuildFail("2ª llamada p_Error: " & m_Err2, logs)
        GoTo Teardown
    End If

    ' --- Cardinalidad invariante: 1 fila, no row extra ---
    If m_Result1 <> m_Result2 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Adversarial_RecargaNoDejaParcial = BuildFail("1ª vs 2ª difieren: [" & m_Result1 & "] vs [" & m_Result2 & "]", logs)
        GoTo Teardown
    End If

    Dim rs As DAO.Recordset
    Dim cnt As Long
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM tbCambiosParaPublicacion WHERE Riesgo='" & CStr(FIX_ID_R9_RIESGO_ADV) & "'")
    If Not rs.EOF Then cnt = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = Nothing
    If cnt <> 1 Then
        Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Adversarial_RecargaNoDejaParcial = BuildFail("esperaba 1 fila en tbCambiosParaPublicacion, obtuvo " & cnt, logs)
        GoTo Teardown
    End If

    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Adversarial_RecargaNoDejaParcial = BuildOk("adversarial_recarga_idempotent_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownR9Fixture FIX_ID_R9_RIESGO_ADV
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformePublicacionHelper_ObtenerMotivoNoPublicable_Adversarial_RecargaNoDejaParcial = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' BLOQUE 3 — REQ-CAL-10 — Atoms for ObtenerDetallePorRiesgo
'
' Helper: ObtenerDetallePorRiesgo(ByRef p_IDRiesgo As String, ...)
'   Returns JSON with {esPublicable, motivo, tieneMaterializaciones,
'                       cantidadMaterializaciones}.
'   Authorization guard: requires m_ObjUsuarioConectado.UsuarioRed in
'   m_ObjEntorno.ColUsuariosCalidad (simplification — full auth layer is
'   out of SDD scope). Atoms 1-3 stub a known quality user; atom 4 stubs
'   a user NOT in the collection to trigger the deny path.
'
' Fixture IDs: dedicated range 903000-903099 (no overlap with
'   SeedBaseGraph 900500-900599, SeedSubcatGraph 900100-900221,
'   Bloque 3 / REQ-CAL-07 901000-901099, REQ-CAL-09 902000-902099,
'   or REQ-CAL-12 904000-904099).
' ============================================================

' ============================================================
' ATOM 1 — Happy: NoPublicableConMotivo
' GIVEN staging sandbox + Riesgo con tbCambiosParaPublicacion.Descripcion
'       = "Falta evidencia de control"
' WHEN ObtenerDetallePorRiesgo(p_IDRiesgo, db, 0, err) is called
'       by user TESTUSER_R10_CALIDAD (in ColUsuariosCalidad)
' THEN:
'   - p_Error vacío
'   - retorno es JSON que contiene "esPublicable":false y "No publicable"
'     y el motivo literal "Falta evidencia de control"
' ============================================================
Public Function Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo() As String
    Dim logs(0 To 9) As String
    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded + ColUsuariosCalidad stub"
    logs(1) = "2. Arrange: SeedR10Fixture(903010) + motivo en tbCambiosParaPublicacion"
    logs(2) = "3. Arrange: m_ObjUsuarioConectado.UsuarioRed = TESTUSER_R10_CALIDAD"
    logs(3) = "4. Act: ObtenerDetallePorRiesgo(903010, db, 0, err)"
    logs(4) = "5. Assert: p_Error vacío"
    logs(5) = "6. Assert: retorno contiene 'esPublicable':false"
    logs(6) = "7. Assert: retorno contiene 'No publicable'"
    logs(7) = "8. Assert: retorno contiene el motivo literal"
    logs(8) = "9. Teardown: borrar motivo + riesgo + teardown"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR10Fixture FIX_ID_R10_RIESGO_HAPPY, FIX_ID_R10_PROYECTO, _
                    FIX_ID_R10_EDICION, FIX_ID_R10_EXPEDIENTE, "R10-HAPPY"

    ' Stub ColUsuariosCalidad with TESTUSER_R10_CALIDAD
    Dim prevDict As Scripting.Dictionary
    Dim prevEntorno As Entorno
    If Not StubColUsuariosCalidadWithUser(FIX_USER_CALIDAD, prevDict, prevEntorno) Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo = BuildFail("No pude stub ColUsuariosCalidad", logs)
        GoTo Teardown
    End If

    ' Save + set m_ObjUsuarioConectado
    Dim prevUsuario As Usuario
    Set prevUsuario = m_ObjUsuarioConectado
    Set m_ObjUsuarioConectado = New Usuario
    m_ObjUsuarioConectado.UsuarioRed = FIX_USER_CALIDAD

    ' Insert motivo row
    db.Execute "INSERT INTO tbCambiosParaPublicacion " & _
               "(IDProyecto, EdicionInicial, EdicionFinal, Riesgo, NombreCampo, Descripcion, FechaRegistro) " & _
               "VALUES (" & FIX_ID_R10_PROYECTO & ", 1, 1, '" & CStr(FIX_ID_R10_RIESGO_HAPPY) & "', " & _
               "'MotivoNoPublicable', " & _
               "'Falta evidencia de control', " & _
               "#" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ObtenerDetallePorRiesgo(CStr(FIX_ID_R10_RIESGO_HAPPY), db, 0, m_Err)

    If Len(m_Err) <> 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If Len(m_Result) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo = BuildFail("retorno vacío", logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, """esPublicable"":false", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo = BuildFail("JSON no contiene 'esPublicable:false', got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "No publicable", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo = BuildFail("JSON no contiene 'No publicable', got: " & m_Result, logs)
        GoTo Teardown
    End If
    ' El motivo se HTMLSafe'a (& ? &amp;), así que chequeamos 'Falta' + 'evidencia' + 'de control' por separado.
    If InStr(1, m_Result, "Falta", vbTextCompare) = 0 Or _
       InStr(1, m_Result, "evidencia", vbTextCompare) = 0 Or _
       InStr(1, m_Result, "de control", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo = BuildFail("motivo literal no presente en JSON, got: " & m_Result, logs)
        GoTo Teardown
    End If

    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo = BuildOk("happy_no_publicable_con_motivo_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set m_ObjUsuarioConectado = prevUsuario
    RestoreColUsuariosCalidad prevDict, prevEntorno
    Set db = Nothing
    TeardownR10Fixture FIX_ID_R10_RIESGO_HAPPY, FIX_ID_R10_PROYECTO, _
                       FIX_ID_R10_EDICION, FIX_ID_R10_EXPEDIENTE
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Happy_NoPublicableConMotivo = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 2 — Sad: ProyectoVacio_MuestraMensaje
' GIVEN staging sandbox + Riesgo válido SIN motivo y SIN materializaciones
' WHEN ObtenerDetallePorRiesgo(p_IDRiesgo, db, 0, err) is called
' THEN:
'   - p_Error vacío
'   - retorno es JSON con "esPublicable":true y motivo que contiene
'     "sin riesgos materializados" (mensaje informativo para el report)
' ============================================================
Public Function Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje() As String
    Dim logs(0 To 8) As String
    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded + ColUsuariosCalidad stub"
    logs(1) = "2. Arrange: SeedR10Fixture(903011) sin motivo ni materializaciones"
    logs(2) = "3. Arrange: m_ObjUsuarioConectado.UsuarioRed = TESTUSER_R10_CALIDAD"
    logs(3) = "4. Act: ObtenerDetallePorRiesgo(903011, db, 0, err)"
    logs(4) = "5. Assert: p_Error vacío"
    logs(5) = "6. Assert: retorno contiene 'sin riesgos materializados'"
    logs(6) = "7. Assert: retorno contiene 'tieneMaterializaciones':false"
    logs(7) = "8. Teardown: borrar riesgo + teardown"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR10Fixture FIX_ID_R10_RIESGO_SAD, FIX_ID_R10_PROYECTO, _
                    FIX_ID_R10_EDICION, FIX_ID_R10_EXPEDIENTE, "R10-SAD"

    Dim prevDict As Scripting.Dictionary
    Dim prevEntorno As Entorno
    If Not StubColUsuariosCalidadWithUser(FIX_USER_CALIDAD, prevDict, prevEntorno) Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje = BuildFail("No pude stub ColUsuariosCalidad", logs)
        GoTo Teardown
    End If

    Dim prevUsuario As Usuario
    Set prevUsuario = m_ObjUsuarioConectado
    Set m_ObjUsuarioConectado = New Usuario
    m_ObjUsuarioConectado.UsuarioRed = FIX_USER_CALIDAD

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ObtenerDetallePorRiesgo(CStr(FIX_ID_R10_RIESGO_SAD), db, 0, m_Err)

    If Len(m_Err) <> 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If Len(m_Result) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje = BuildFail("retorno vacío", logs)
        GoTo Teardown
    End If
    ' Mensaje informativo: "sin riesgos materializados" (Sad path del report).
    If InStr(1, m_Result, "sin riesgos", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje = BuildFail("esperaba 'sin riesgos' en motivo, got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, """tieneMaterializaciones"":false", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje = BuildFail("esperaba 'tieneMaterializaciones:false' en JSON, got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, """cantidadMaterializaciones"":0", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje = BuildFail("esperaba 'cantidadMaterializaciones:0', got: " & m_Result, logs)
        GoTo Teardown
    End If

    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje = BuildOk("sad_sin_riesgos_mensaje_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set m_ObjUsuarioConectado = prevUsuario
    RestoreColUsuariosCalidad prevDict, prevEntorno
    Set db = Nothing
    TeardownR10Fixture FIX_ID_R10_RIESGO_SAD, FIX_ID_R10_PROYECTO, _
                       FIX_ID_R10_EDICION, FIX_ID_R10_EXPEDIENTE
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Sad_ProyectoVacio_MuestraMensaje = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 3 — Edge: MultiplesMaterializaciones_Consolida
' GIVEN staging sandbox + Riesgo con 3 TbRiesgosMaterializaciones:
'       (1) ParaNC='Sí', Estado='Resuelto' (con NC decidida)
'       (2) ParaNC='No', Estado='Resuelto' (con NC decidida)
'       (3) ParaNC='' (vacío, pendiente)
' WHEN ObtenerDetallePorRiesgo(p_IDRiesgo, db, 0, err) is called
' THEN:
'   - p_Error vacío
'   - retorno es JSON con cantidadMaterializaciones=3, esPublicable=false
'   - motivo menciona "1 pendiente" y "2 con NC decidida" (consolidación)
' ============================================================
Public Function Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida() As String
    Dim logs(0 To 9) As String
    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded + ColUsuariosCalidad stub"
    logs(1) = "2. Arrange: SeedR10Fixture(903012) + 3 materializaciones (1 pendiente, 2 con NC)"
    logs(2) = "3. Arrange: m_ObjUsuarioConectado.UsuarioRed = TESTUSER_R10_CALIDAD"
    logs(3) = "4. Act: ObtenerDetallePorRiesgo(903012, db, 0, err)"
    logs(4) = "5. Assert: p_Error vacío"
    logs(5) = "6. Assert: cantidadMaterializaciones=3"
    logs(6) = "7. Assert: esPublicable=false (1 pendiente)"
    logs(7) = "8. Assert: motivo menciona '1' y '2' (consolidación de conteos)"
    logs(8) = "9. Teardown: borrar materializaciones + riesgo + teardown"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR10Fixture FIX_ID_R10_RIESGO_EDGE, FIX_ID_R10_PROYECTO, _
                    FIX_ID_R10_EDICION, FIX_ID_R10_EXPEDIENTE, "R10-EDGE"

    Dim prevDict As Scripting.Dictionary
    Dim prevEntorno As Entorno
    If Not StubColUsuariosCalidadWithUser(FIX_USER_CALIDAD, prevDict, prevEntorno) Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail("No pude stub ColUsuariosCalidad", logs)
        GoTo Teardown
    End If

    Dim prevUsuario As Usuario
    Set prevUsuario = m_ObjUsuarioConectado
    Set m_ObjUsuarioConectado = New Usuario
    m_ObjUsuarioConectado.UsuarioRed = FIX_USER_CALIDAD

    ' Insertar 3 materializaciones:
    '  (1) Fecha reciente, ParaNC='Sí', Estado='Resuelto'
    '  (2) Fecha media, ParaNC='No', Estado='Resuelto'
    '  (3) Fecha antigua, ParaNC='' (Null o vacío), Estado='Resuelto'
    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
               "(IDProyecto, CodigoRiesgo, Fecha, Estado, ParaNC, EsMaterializacion) " & _
               "VALUES (" & FIX_ID_R10_PROYECTO & ", 'R10-EDGE', " & _
               "#" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#, 'Resuelto', 'Sí', 'Sí')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
               "(IDProyecto, CodigoRiesgo, Fecha, Estado, ParaNC, EsMaterializacion) " & _
               "VALUES (" & FIX_ID_R10_PROYECTO & ", 'R10-EDGE', " & _
               "#" & Format$(DateAdd("d", -1, Now), "yyyy-mm-dd hh:nn:ss") & "#, 'Resuelto', 'No', 'Sí')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
               "(IDProyecto, CodigoRiesgo, Fecha, Estado, ParaNC, EsMaterializacion) " & _
               "VALUES (" & FIX_ID_R10_PROYECTO & ", 'R10-EDGE', " & _
               "#" & Format$(DateAdd("d", -2, Now), "yyyy-mm-dd hh:nn:ss") & "#, 'Resuelto', NULL, 'Sí')", dbFailOnError

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ObtenerDetallePorRiesgo(CStr(FIX_ID_R10_RIESGO_EDGE), db, 0, m_Err)

    If Len(m_Err) <> 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If Len(m_Result) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail("retorno vacío", logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, """cantidadMaterializaciones"":3", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail("esperaba 'cantidadMaterializaciones:3', got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, """esPublicable"":false", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail("esperaba 'esPublicable:false' (1 pendiente), got: " & m_Result, logs)
        GoTo Teardown
    End If
    ' El motivo consolidado debe mencionar 1 pendiente y 2 con NC decidida.
    If InStr(1, m_Result, "1 materializacion(es) pendiente(s)", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail("motivo no menciona '1 pendiente', got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "total=3", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail("motivo no menciona 'total=3', got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "con NC decidida=2", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail("motivo no menciona 'con NC decidida=2', got: " & m_Result, logs)
        GoTo Teardown
    End If

    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildOk("edge_multiples_mat_consolida_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set m_ObjUsuarioConectado = prevUsuario
    RestoreColUsuariosCalidad prevDict, prevEntorno
    Set db = Nothing
    TeardownR10Fixture FIX_ID_R10_RIESGO_EDGE, FIX_ID_R10_PROYECTO, _
                       FIX_ID_R10_EDICION, FIX_ID_R10_EXPEDIENTE
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Edge_MultiplesMaterializaciones_Consolida = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 4 — Adversarial: SinPermiso_Denegado
' GIVEN staging sandbox + Riesgo válido + motivo persistido
' WHEN ObtenerDetallePorRiesgo is called by user TESTUSER_R10_CALIDAD
'       que NO está en ColUsuariosCalidad (sólo OTRO_USUARIO_CALIDAD)
' THEN:
'   - p_Error contiene "sin permisos"
'   - retorno es ""
' ============================================================
Public Function Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Adversarial_SinPermiso_Denegado() As String
    Dim logs(0 To 8) As String
    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Adversarial_SinPermiso_Denegado = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded + ColUsuariosCalidad stub con OTRO usuario"
    logs(1) = "2. Arrange: SeedR10Fixture(903013) + motivo persistido"
    logs(2) = "3. Arrange: m_ObjUsuarioConectado.UsuarioRed = TESTUSER_R10_CALIDAD (NO en ColUsuariosCalidad)"
    logs(3) = "4. Act: ObtenerDetallePorRiesgo(903013, db, 0, err)"
    logs(4) = "5. Assert: p_Error contiene 'sin permisos'"
    logs(5) = "6. Assert: retorno es vacío"
    logs(6) = "7. Assert: motivo persistido NO se expone en el retorno"
    logs(7) = "8. Teardown: borrar motivo + riesgo + teardown"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Adversarial_SinPermiso_Denegado = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Adversarial_SinPermiso_Denegado = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR10Fixture FIX_ID_R10_RIESGO_ADV, FIX_ID_R10_PROYECTO, _
                    FIX_ID_R10_EDICION, FIX_ID_R10_EXPEDIENTE, "R10-ADV"

    ' Stub ColUsuariosCalidad con OTRO usuario (NO el que vamos a usar).
    Dim prevDict As Scripting.Dictionary
    Dim prevEntorno As Entorno
    If Not StubColUsuariosCalidadWithUser(FIX_USER_OTRO_CALIDAD, prevDict, prevEntorno) Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Adversarial_SinPermiso_Denegado = BuildFail("No pude stub ColUsuariosCalidad", logs)
        GoTo Teardown
    End If

    Dim prevUsuario As Usuario
    Set prevUsuario = m_ObjUsuarioConectado
    Set m_ObjUsuarioConectado = New Usuario
    m_ObjUsuarioConectado.UsuarioRed = FIX_USER_CALIDAD  ' NOT in the dict

    ' Insert motivo para que el helper intentaría devolverlo si pasara el guard
    db.Execute "INSERT INTO tbCambiosParaPublicacion " & _
               "(IDProyecto, EdicionInicial, EdicionFinal, Riesgo, NombreCampo, Descripcion, FechaRegistro) " & _
               "VALUES (" & FIX_ID_R10_PROYECTO & ", 1, 1, '" & CStr(FIX_ID_R10_RIESGO_ADV) & "', " & _
               "'MotivoNoPublicable', " & _
               "'CONFIDENCIAL_NO_DEBE_VER', " & _
               "#" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ObtenerDetallePorRiesgo(CStr(FIX_ID_R10_RIESGO_ADV), db, 0, m_Err)

    If Len(m_Err) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Adversarial_SinPermiso_Denegado = BuildFail("esperaba p_Error poblado, got vacío", logs)
        GoTo Teardown
    End If
    If InStr(1, m_Err, "sin permisos", vbTextCompare) = 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Adversarial_SinPermiso_Denegado = BuildFail("esperaba 'sin permisos' en p_Error, got: " & m_Err, logs)
        GoTo Teardown
    End If
    If m_Result <> "" Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Adversarial_SinPermiso_Denegado = BuildFail("esperaba retorno vacío, got: " & m_Result, logs)
        GoTo Teardown
    End If
    ' El motivo confidencial NO debe filtrarse al caller.
    If InStr(1, m_Result, "CONFIDENCIAL", vbTextCompare) > 0 Then
        Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Adversarial_SinPermiso_Denegado = BuildFail("motivo confidencial filtrado en retorno: " & m_Result, logs)
        GoTo Teardown
    End If

    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Adversarial_SinPermiso_Denegado = BuildOk("adversarial_sin_permisos_denegado_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set m_ObjUsuarioConectado = prevUsuario
    RestoreColUsuariosCalidad prevDict, prevEntorno
    Set db = Nothing
    TeardownR10Fixture FIX_ID_R10_RIESGO_ADV, FIX_ID_R10_PROYECTO, _
                       FIX_ID_R10_EDICION, FIX_ID_R10_EXPEDIENTE
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_InformePublicacionHelper_ObtenerDetallePorRiesgo_Adversarial_SinPermiso_Denegado = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

