Attribute VB_Name = "TestFixtures"
Option Compare Database
Option Explicit

Private Const TEST_ID_BASE As Long = 994000000
Private Const E2E_EXPORT_TEST_ID_BASE As Long = 994012000

Public Function CacheFixtureBaseId() As Long
    CacheFixtureBaseId = TEST_ID_BASE
End Function

Public Function E2EExportFixtureBaseId() As Long
    E2EExportFixtureBaseId = E2E_EXPORT_TEST_ID_BASE
End Function

Public Function SeedExpedienteLugarFixture(ByVal p_IDExp As Long, ByVal p_IDLugar As Long, ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Db As DAO.Database
    p_Error = ""
    SeedExpedienteLugarFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function

    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "DELETE * FROM TbExpedientesLugaresEjecucion WHERE IDExpediente=" & p_IDExp & " AND IDLugarEjecucion=" & p_IDLugar & ";", dbFailOnError
    m_Db.Execute "DELETE * FROM TbExpedientes WHERE IDExpediente=" & p_IDExp & ";", dbFailOnError

    m_Db.Execute "INSERT INTO TbExpedientes (IDExpediente, CodExp, Nemotecnico) VALUES (" & p_IDExp & ", 'COD-PR3-LUGAR', 'NEMO-PR3-LUGAR');", dbFailOnError
    m_Db.Execute "INSERT INTO TbExpedientesLugaresEjecucion (IDExpedienteLugarEjecucion, IDExpediente, IDLugarEjecucion) VALUES (" & p_IDExp & ", " & p_IDExp & ", " & p_IDLugar & ");", dbFailOnError

    SeedExpedienteLugarFixture = True
    Exit Function

EH:
    p_Error = "SeedExpedienteLugarFixture: " & Err.Description
End Function

Public Function TeardownExpedienteLugarFixture(ByVal p_IDExp As Long, ByVal p_IDLugar As Long, ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Db As DAO.Database
    p_Error = ""
    TeardownExpedienteLugarFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function

    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "DELETE * FROM TbExpedientesLugaresEjecucion WHERE IDExpediente=" & p_IDExp & " AND IDLugarEjecucion=" & p_IDLugar & ";", dbFailOnError
    m_Db.Execute "DELETE * FROM TbExpedientes WHERE IDExpediente=" & p_IDExp & ";", dbFailOnError

    TeardownExpedienteLugarFixture = True
    Exit Function

EH:
    p_Error = "TeardownExpedienteLugarFixture: " & Err.Description
End Function

Public Function SeedExpedienteRacFixture(ByVal p_IDExp As Long, ByVal p_IDRac As Long, ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Db As DAO.Database
    p_Error = ""
    SeedExpedienteRacFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function

    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "DELETE * FROM TbExpedientesRACS WHERE IDExpediente=" & p_IDExp & " AND IDRAC=" & p_IDRac & ";", dbFailOnError
    m_Db.Execute "DELETE * FROM TbExpedientes WHERE IDExpediente=" & p_IDExp & ";", dbFailOnError

    m_Db.Execute "INSERT INTO TbExpedientes (IDExpediente, CodExp, Nemotecnico) VALUES (" & p_IDExp & ", 'COD-PR3-RAC', 'NEMO-PR3-RAC');", dbFailOnError
    m_Db.Execute "INSERT INTO TbExpedientesRACS (IDRACExpediente, IDExpediente, IDRAC) VALUES (" & p_IDExp & ", " & p_IDExp & ", " & p_IDRac & ");", dbFailOnError

    SeedExpedienteRacFixture = True
    Exit Function

EH:
    p_Error = "SeedExpedienteRacFixture: " & Err.Description
End Function

Public Function TeardownExpedienteRacFixture(ByVal p_IDExp As Long, ByVal p_IDRac As Long, ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Db As DAO.Database
    p_Error = ""
    TeardownExpedienteRacFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function

    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "DELETE * FROM TbExpedientesRACS WHERE IDExpediente=" & p_IDExp & " AND IDRAC=" & p_IDRac & ";", dbFailOnError
    m_Db.Execute "DELETE * FROM TbExpedientes WHERE IDExpediente=" & p_IDExp & ";", dbFailOnError

    TeardownExpedienteRacFixture = True
    Exit Function

EH:
    p_Error = "TeardownExpedienteRacFixture: " & Err.Description
End Function

Public Function SetupE2EBatchSchemaSandbox(ByRef p_Error As String) As Boolean
    On Error GoTo EH

    p_Error = ""
    SetupE2EBatchSchemaSandbox = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    If Not MigrateE2EExportTraceabilitySchema(p_Error) Then Exit Function
    If Not TeardownE2EBatchSchemaSandbox(p_Error) Then Exit Function

    SetupE2EBatchSchemaSandbox = True
    Exit Function

EH:
    p_Error = "SetupE2EBatchSchemaSandbox: " & Err.Description
End Function

Public Function TeardownE2EBatchSchemaSandbox(ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Db As DAO.Database
    p_Error = ""
    TeardownE2EBatchSchemaSandbox = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function

    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "DELETE FROM TbE2EExportBatchDetalle " & _
                 "WHERE IDBatch IN (SELECT IDBatch FROM TbE2EExportBatch " & _
                 "WHERE UsuarioConectado LIKE 'qa.user%' OR SessionId LIKE 'S-%');", dbFailOnError
    m_Db.Execute "DELETE FROM TbE2EExportBatch " & _
                 "WHERE UsuarioConectado LIKE 'qa.user%' OR SessionId LIKE 'S-%';", dbFailOnError
    m_Db.Execute "DELETE FROM TbE2EExportSeleccionTemp " & _
                 "WHERE UsuarioConectado LIKE 'qa.user%' OR SessionId LIKE 'S-%';", dbFailOnError
    m_Db.Execute "DELETE FROM TbE2EJsonDestinationUserConfig WHERE UsuarioRed LIKE 'qa.user%';", dbFailOnError

    TeardownE2EBatchSchemaSandbox = True
    Exit Function

EH:
    p_Error = "TeardownE2EBatchSchemaSandbox: " & Err.Description
End Function

Public Function TeardownE2EJsonDestinationConfigSandbox(ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Db As DAO.Database
    p_Error = ""
    TeardownE2EJsonDestinationConfigSandbox = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    If Not EnsureE2EJsonDestinationConfigSchema(p_Error) Then Exit Function

    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "DELETE FROM TbE2EJsonDestinationUserConfig WHERE UsuarioRed LIKE 'qa.user%';", dbFailOnError

    TeardownE2EJsonDestinationConfigSandbox = True
    Exit Function

EH:
    p_Error = "TeardownE2EJsonDestinationConfigSandbox: " & Err.Description
End Function

Public Function SeedE2EExportExpedienteFixture( _
    ByVal p_IDExpediente As Long, _
    ByVal p_OrdinalE2E As Long, _
    ByVal p_HashActual As String, _
    ByVal p_HashUltimaExportacion As String, _
    ByVal p_Marker As String, _
    ByRef p_Error As String) As Boolean

    On Error GoTo EH

    Dim m_Db As DAO.Database
    Dim hashUltimaSql As String

    p_Error = ""
    SeedE2EExportExpedienteFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    If Not EnsureE2EExportTraceabilitySchema(p_Error) Then Exit Function


    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "DELETE FROM TbE2EExportSeleccionTemp WHERE IDExpediente=" & CLng(p_IDExpediente) & ";", dbFailOnError
    m_Db.Execute "DELETE FROM TbE2EExportBatchDetalle WHERE IDExpediente=" & CLng(p_IDExpediente) & ";", dbFailOnError
    m_Db.Execute "DELETE FROM TbExpedientes WHERE (IDExpediente=" & CLng(p_IDExpediente) & _
                 " OR OrdinalE2E=" & CLng(p_OrdinalE2E) & _
                 " OR Nemotecnico='" & SqlStr(p_Marker) & "');", dbFailOnError

    If Trim$(p_HashUltimaExportacion) = "" Then
        hashUltimaSql = "Null"
    Else
        hashUltimaSql = "'" & SqlStr(p_HashUltimaExportacion) & "'"
    End If

    m_Db.Execute "INSERT INTO TbExpedientes (" & _
                 "IDExpediente, OrdinalE2E, Nemotecnico, Titulo, HashActual, HashUltimaExportacion" & _
                 ") VALUES (" & CLng(p_IDExpediente) & ", " & CLng(p_OrdinalE2E) & _
                 ", '" & SqlStr(p_Marker) & "', 'Fixture E2E export " & SqlStr(p_Marker) & _
                 "', '" & SqlStr(p_HashActual) & "', " & hashUltimaSql & ");", dbFailOnError

    SeedE2EExportExpedienteFixture = True
    Exit Function

EH:
    p_Error = "SeedE2EExportExpedienteFixture: " & Err.Description
End Function

Public Function TeardownE2EExportFixture(ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Db As DAO.Database
    p_Error = ""
    TeardownE2EExportFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    If Not EnsureE2EExportTraceabilitySchema(p_Error) Then Exit Function

    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "DELETE FROM TbE2EExportBatchDetalle WHERE IDBatch IN (" & _
                 "SELECT IDBatch FROM TbE2EExportBatch WHERE Left(Nz(UsuarioConectado,''),9)='qa.export' OR Left(Nz(SessionId,''),9)='S-EXPORT-');", dbFailOnError
    m_Db.Execute "DELETE FROM TbE2EExportBatch WHERE Left(Nz(UsuarioConectado,''),9)='qa.export' OR Left(Nz(SessionId,''),9)='S-EXPORT-';", dbFailOnError
    m_Db.Execute "DELETE FROM TbE2EExportSeleccionTemp WHERE Left(Nz(UsuarioConectado,''),9)='qa.export' OR Left(Nz(SessionId,''),9)='S-EXPORT-';", dbFailOnError
    m_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente BETWEEN " & E2E_EXPORT_TEST_ID_BASE & _
                 " AND " & (E2E_EXPORT_TEST_ID_BASE + 999) & ";", dbFailOnError

    TeardownE2EExportFixture = True
    Exit Function

EH:
    p_Error = "TeardownE2EExportFixture: " & Err.Description
End Function

Public Function CreateE2EExportBatchFixture( _
    ByVal p_UsuarioConectado As String, _
    ByVal p_SessionId As String, _
    ByRef p_IDBatch As Long, _
    ByRef p_Error As String) As Boolean

    On Error GoTo EH

    Dim m_Db As DAO.Database
    Dim rs As DAO.Recordset

    p_Error = ""
    p_IDBatch = 0
    CreateE2EExportBatchFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    If Not EnsureE2EBatchManagementSchema(p_Error) Then Exit Function

    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "INSERT INTO TbE2EExportBatch (SessionId, UsuarioConectado, Estado, CreatedAt, TotalSeleccionados, TotalExportados) VALUES (" & _
                 "'" & SqlStr(p_SessionId) & "', '" & SqlStr(p_UsuarioConectado) & "', 'CREATED', Now(), 0, 0);", dbFailOnError

    Set rs = m_Db.OpenRecordset("SELECT @@IDENTITY AS NewId", dbOpenSnapshot)
    If Not rs.EOF Then p_IDBatch = CLng(Nz(rs.Fields("NewId").value, 0))
    rs.Close
    Set rs = Nothing

    CreateE2EExportBatchFixture = (p_IDBatch > 0)
    If Not CreateE2EExportBatchFixture Then p_Error = "CreateE2EExportBatchFixture: no IDBatch returned"
    Exit Function

EH:
    p_Error = "CreateE2EExportBatchFixture: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Public Function SeedAutosaveExpedienteFixture(ByVal p_IDExp As Long, ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Db As DAO.Database
    p_Error = ""
    SeedAutosaveExpedienteFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function

    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    If Not TeardownAutosaveExpedienteFixture(p_IDExp, p_Error) Then Exit Function

    m_Db.Execute "INSERT INTO TbExpedientes (" & _
                 "IDExpediente, CodExp, Nemotecnico, Titulo, Ambito, FECHAINICIOLICITACION, FECHAADJUDICACION" & _
                 ") VALUES (" & p_IDExp & _
                 ", 'TEST-AUTOSAVE-GF', 'TEST-AUTOSAVE-GF', 'Fixture autosave General/Fechas', 'TEST', #2026-01-10#, #2026-02-10#);", dbFailOnError

    SeedAutosaveExpedienteFixture = True
    Exit Function

EH:
    p_Error = "SeedAutosaveExpedienteFixture: " & Err.Description
End Function

Public Function SeedAutosaveHitoFixture(ByVal p_IDExp As Long, ByVal p_IDHito As Long, ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Db As DAO.Database
    p_Error = ""
    SeedAutosaveHitoFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "DELETE * FROM TbExpedientesHitos WHERE IDHitoExpediente=" & p_IDHito & ";", dbFailOnError
    m_Db.Execute "INSERT INTO TbExpedientesHitos (IDHitoExpediente, IDExpediente, Descripcion, FechaHito, FechaGarantiaHito, Importe) VALUES (" & _
                 p_IDHito & ", " & p_IDExp & ", 'TEST-AUTOSAVE-HITO', #2026-03-01#, #2026-04-01#, 123.45);", dbFailOnError

    SeedAutosaveHitoFixture = True
    Exit Function

EH:
    p_Error = "SeedAutosaveHitoFixture: " & Err.Description
End Function

Public Function SeedAutosaveModificadoFixture(ByVal p_IDExp As Long, ByVal p_IDModificado As Long, ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Db As DAO.Database
    p_Error = ""
    SeedAutosaveModificadoFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "DELETE * FROM TbExpedientesModificados WHERE IDExpedienteModificado=" & p_IDModificado & ";", dbFailOnError
    m_Db.Execute "INSERT INTO TbExpedientesModificados (IDExpedienteModificado, IDExpediente, NModificado, FechaFirmaModificado, FechaFinModificado, Descripcion) VALUES (" & _
                 p_IDModificado & ", " & p_IDExp & ", 'TEST-AUTOSAVE-MOD', #2026-05-01#, #2026-06-01#, 'Fixture autosave modificado');", dbFailOnError

    SeedAutosaveModificadoFixture = True
    Exit Function

EH:
    p_Error = "SeedAutosaveModificadoFixture: " & Err.Description
End Function

Public Function TeardownAutosaveExpedienteFixture(ByVal p_IDExp As Long, ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Db As DAO.Database
    p_Error = ""
    TeardownAutosaveExpedienteFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    Set m_Db = getdb(p_Error)
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    m_Db.Execute "DELETE * FROM TbExpedientesHitos WHERE IDExpediente=" & p_IDExp & ";", dbFailOnError
    m_Db.Execute "DELETE * FROM TbExpedientesModificados WHERE IDExpediente=" & p_IDExp & ";", dbFailOnError
    m_Db.Execute "DELETE * FROM TbExpedientes WHERE IDExpediente=" & p_IDExp & " AND CodExp='TEST-AUTOSAVE-GF';", dbFailOnError

    TeardownAutosaveExpedienteFixture = True
    Exit Function

EH:
    p_Error = "TeardownAutosaveExpedienteFixture: " & Err.Description
End Function

Public Function SetupE2ESelectionFixture( _
    ByVal p_Usuario As String, _
    ByVal p_SessionId As String, _
    ByVal p_IdsCsv As String, _
    ByRef p_Error As String, _
    Optional ByVal p_Db As DAO.Database = Nothing) As Boolean

    On Error GoTo EH

    Dim m_Db As DAO.Database
    Dim m_Ids As Collection
    Dim m_Item As Variant

    p_Error = ""
    SetupE2ESelectionFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    If Not EnsureE2EBatchManagementSchema(p_Error) Then Exit Function

    If p_Db Is Nothing Then
        Set m_Db = getdb(p_Error)
    Else
        Set m_Db = p_Db
    End If
    If m_Db Is Nothing Then
        If p_Error = "" Then p_Error = "No se pudo abrir backend sandbox"
        Exit Function
    End If

    If Not ClearSelectionTempRows(p_Usuario, p_SessionId, p_Error, m_Db) Then Exit Function

    Set m_Ids = ParseCsvToCollection(p_IdsCsv)
    For Each m_Item In m_Ids
        If Not AddSelectionTempRow(p_Usuario, p_SessionId, CLng(m_Item), p_Error, m_Db) Then Exit Function
    Next m_Item

    SetupE2ESelectionFixture = True
    Exit Function

EH:
    p_Error = "SetupE2ESelectionFixture: " & Err.Description
End Function

Public Function ParseCsvToCollection(ByVal p_Csv As String) As Collection
    Dim parts() As String
    Dim i As Long
    Dim value As String
    Dim c As Collection

    Set c = New Collection
    If Trim$(p_Csv) = "" Then
        Set ParseCsvToCollection = c
        Exit Function
    End If

    parts = Split(p_Csv, ",")
    For i = LBound(parts) To UBound(parts)
        value = Trim$(parts(i))
        If value <> "" Then c.Add value
    Next i
    Set ParseCsvToCollection = c
End Function

