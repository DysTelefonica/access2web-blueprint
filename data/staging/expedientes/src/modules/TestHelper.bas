Attribute VB_Name = "TestHelper"
Option Compare Database
Option Explicit

' Canonical JSON helpers for VBA test contract

Public Function BuildJsonOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildJsonOk = "{""ok"":true,""value"":" & JsonValue(value) & ",""payload"":null,""error"":null,""logs"":" & JsonStringArray(logs) & "}"
End Function

Public Function BuildJsonFail(ByVal errorMsg As String, ByRef logs() As String) As String
    BuildJsonFail = "{""ok"":false,""value"":null,""payload"":null,""error"":""" & EscapeJsonString(errorMsg) & """,""logs"":" & JsonStringArray(logs) & "}"
End Function

Public Function EscapeJsonString(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, Chr$(34), "\" & Chr$(34))
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJsonString = s
End Function

Public Function JsonStringArray(ByRef logs() As String) As String
    On Error GoTo EmptyLogs

    Dim i As Long
    Dim parts As String

    For i = LBound(logs) To UBound(logs)
        If Len(logs(i)) > 0 Then
            If Len(parts) > 0 Then parts = parts & ","
            parts = parts & """" & EscapeJsonString(logs(i)) & """"
        End If
    Next i

    JsonStringArray = "[" & parts & "]"
    Exit Function

EmptyLogs:
    JsonStringArray = "[]"
End Function

Public Function SqlStr(ByVal value As String) As String
    SqlStr = Replace(value, "'", "''")
End Function

Public Function ForceLocalBackend(ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim m_Rst As DAO.Recordset
    Dim m_SandboxPath As String
    Dim m_Password As String
    Dim m_Db As DAO.Database

    p_Error = ""
    ForceLocalBackend = False

    Set m_Rst = CurrentDb.OpenRecordset("SELECT BackendSandbox, PasswordBackend FROM TbConfiguracionBackends ORDER BY Id;")
    If m_Rst.EOF Then
        p_Error = "No se encontró configuración de backend para tests"
        GoTo Cleanup
    End If

    m_SandboxPath = Trim$(Nz(m_Rst!BackendSandbox, ""))
    m_Password = Nz(m_Rst!PasswordBackend, "")

    If m_SandboxPath = "" Then
        p_Error = "BackendSandbox vacío en TbConfiguracionBackends"
        GoTo Cleanup
    End If

    If Not fso.FileExists(m_SandboxPath) Then
        p_Error = "Sandbox no alcanzable: " & m_SandboxPath
        GoTo Cleanup
    End If

    On Error GoTo EH
    Set m_Db = DBEngine(0).OpenDatabase(m_SandboxPath, False, True, ";pwd=" & m_Password)
    m_Db.Close
    Set m_Db = Nothing

    CloseCachedBackendConnection
    TestOnlyConfigureTestingBackend m_SandboxPath, m_Password
    m_TestingMode = True

    ForceLocalBackend = True

Cleanup:
    On Error Resume Next
    If Not m_Rst Is Nothing Then m_Rst.Close
    Set m_Rst = Nothing
    Set m_Db = Nothing
    Exit Function

EH:
    If p_Error = "" Then p_Error = "ForceLocalBackend: " & Err.Description
    Resume Cleanup
End Function

Public Function AssertSandboxBackend(Optional ByRef p_Error As String) As Boolean
    p_Error = ""
    AssertSandboxBackend = False

    If Not m_TestingMode Then
        p_Error = "Sandbox guard: m_TestingMode=False"
        Exit Function
    End If

    If Trim$(TestOnlyGetTestingBackendURL()) = "" Then
        p_Error = "Sandbox guard: backend sandbox no configurado"
        Exit Function
    End If

    If Not fso.FileExists(TestOnlyGetTestingBackendURL()) Then
        p_Error = "Sandbox guard: backend sandbox no alcanzable"
        Exit Function
    End If

    If InStr(1, UCase$(TestOnlyGetTestingBackendURL()), "PROD", vbTextCompare) > 0 Then
        p_Error = "Sandbox guard: ruta sospechosa de PROD"
        Exit Function
    End If

    ' Detect stale cached connection — if m_dbCached points to a different
    ' backend than the currently configured testing URL, the guard fails.
    If TestOnlyIsSandboxCacheStale() Then
        p_Error = "Sandbox guard: cached connection is stale (points to different backend)"
        Exit Function
    End If

    AssertSandboxBackend = True
End Function

Public Function EnsureSandboxBackend(ByRef p_Error As String) As Boolean
    p_Error = ""

    If Not m_TestingMode Or Trim$(TestOnlyGetTestingBackendURL()) = "" Then
        Call ForceLocalBackend(p_Error)
        If p_Error <> "" Then
            EnsureSandboxBackend = False
            Exit Function
        End If
    End If

    EnsureSandboxBackend = AssertSandboxBackend(p_Error)
    If Not EnsureSandboxBackend Then
        If InStr(1, p_Error, "cached connection is stale", vbTextCompare) > 0 Then
            CloseCachedBackendConnection
            EnsureSandboxBackend = ForceLocalBackend(p_Error)
        End If
    End If
End Function

Public Sub ResetTestSession()
    On Error Resume Next
    m_TestingMode = False
    TestOnlyClearTestingBackend
    CloseCachedBackendConnection
    TestOnlyResetBackendConfigOverride
    Application.TempVars.Remove "BackendPathSandbox"
    Application.TempVars.Remove "BackendPathConfigurado"
    Application.TempVars.Remove "DatosEnLocal"
End Sub

Public Function BeginTestSession(ByRef p_Error As String) As Boolean
    p_Error = ""
    BeginTestSession = EnsureSandboxBackend(p_Error)
End Function

Public Function EndTestSession(Optional ByRef p_Error As String) As Boolean
    p_Error = ""
    ResetTestSession
    EndTestSession = True
End Function

Public Function GetTestDb(Optional ByRef p_Error As String) As DAO.Database
    p_Error = ""
    If Not EnsureSandboxBackend(p_Error) Then
        Set GetTestDb = Nothing
        Exit Function
    End If
    Set GetTestDb = getdb(p_Error)
    If GetTestDb Is Nothing And p_Error = "" Then p_Error = "GetTestDb: getdb returned Nothing"
End Function

Public Function EnsureWorkingTableClean( _
    ByVal p_Db As DAO.Database, _
    ByVal p_TableName As String, _
    ByVal p_WhereClause As String, _
    Optional ByRef p_Error As String) As Boolean

    On Error GoTo EH

    p_Error = ""
    EnsureWorkingTableClean = False

    If p_Db Is Nothing Then
        p_Error = "EnsureWorkingTableClean: explicit DAO.Database is required"
        Exit Function
    End If
    If Trim$(p_TableName) = "" Then
        p_Error = "EnsureWorkingTableClean: table name is required"
        Exit Function
    End If
    If Trim$(p_WhereClause) = "" Then
        p_Error = "EnsureWorkingTableClean: guarded WHERE clause is required"
        Exit Function
    End If

    p_Db.Execute "DELETE FROM " & p_TableName & " WHERE " & p_WhereClause, dbFailOnError
    EnsureWorkingTableClean = True
    Exit Function

EH:
    p_Error = "EnsureWorkingTableClean: " & Err.Description
End Function

' PRUEBA-003/REFAC-4b2b — Idempotent fixture for the supplier-chain tree
' (closes #60). Inserts a Collection<ExpedienteSuministrador> into
' TbExpedientesSuministradores using the injected sandboxDb. Idempotency
' is enforced by a per-row PK check on IDExpedienteSuministrador — a row
' that already exists in the target db is skipped, NOT updated. This keeps
' the fixture safe to call across multiple test atoms without re-seeding.
' Returns the number of rows actually inserted (0..p_Items.Count).
Public Function SeedArbolSuministradores( _
    ByVal p_Items As Collection, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String _
) As Long
    Dim dbToUse As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlCheck As String
    Dim sqlInsert As String
    Dim item As ExpedienteSuministrador
    Dim vIdPadre As String
    Dim vDescripcon As String

    On Error GoTo EH

    p_Error = ""
    SeedArbolSuministradores = 0

    If p_Db Is Nothing Then
        Set dbToUse = getdb()
    Else
        Set dbToUse = p_Db
    End If

    If p_Items Is Nothing Then Exit Function

    Dim v As Variant
    For Each v In p_Items
        Set item = Nothing
        If TypeName(v) = "ExpedienteSuministrador" Then
            Set item = v
        End If
        If item Is Nothing Then
            p_Error = "SeedArbolSuministradores: collection element is not ExpedienteSuministrador"
            Err.Raise 1000
        End If

        Dim sIDExpSum As String
        sIDExpSum = item.getPropiedad("IDExpedienteSuministrador", p_Error)
        If p_Error <> "" Then Err.Raise 1000

        ' Idempotency: skip if PK already present in the target table
        sqlCheck = "SELECT COUNT(*) AS n FROM TbExpedientesSuministradores " & _
                   "WHERE IDExpedienteSuministrador=" & sIDExpSum
        Set rs = dbToUse.OpenRecordset(sqlCheck, dbOpenSnapshot)
        If CLng(rs!n) > 0 Then
            rs.Close
            Set rs = Nothing
        Else
            rs.Close
            Set rs = Nothing

            Dim sIdPadreVal As String
            Dim vPadre As Variant
            vPadre = item.getPropiedad("IdPadre", p_Error)
            If p_Error <> "" Then Err.Raise 1000
            If IsNull(vPadre) Or vPadre = "" Or vPadre = 0 Then
                sIdPadreVal = "Null"
            Else
                sIdPadreVal = CStr(vPadre)
            End If

            vDescripcon = CStr(Nz(item.getPropiedad("Descripcon", p_Error), ""))
            If p_Error <> "" Then Err.Raise 1000

            sqlInsert = "INSERT INTO TbExpedientesSuministradores (" & _
                        "IDExpedienteSuministrador, IDExpediente, IDSuministrador, " & _
                        "IdPadre, ContratistaPrincipal, SubContratista, Descripcon) " & _
                        "VALUES (" & _
                        sIDExpSum & ", " & _
                        item.getPropiedad("IDExpediente", p_Error) & ", " & _
                        item.getPropiedad("IDSuministrador", p_Error) & ", " & _
                        sIdPadreVal & ", " & _
                        "'" & SqlStr(CStr(item.getPropiedad("ContratistaPrincipal", p_Error))) & "', " & _
                        "'" & SqlStr(CStr(item.getPropiedad("SubContratista", p_Error))) & "', " & _
                        "'" & SqlStr(vDescripcon) & "')"
            If p_Error <> "" Then Err.Raise 1000

            dbToUse.Execute sqlInsert, dbFailOnError
            SeedArbolSuministradores = SeedArbolSuministradores + 1
        End If
    Next v

    Exit Function

EH:
    If Err.Number <> 1000 Then p_Error = "SeedArbolSuministradores: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Function JsonValue(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        JsonValue = "null"
        Exit Function
    End If

    Select Case VarType(value)
        Case vbBoolean
            ' Canonical JSON literals (not "verdadero"/"falso"): CStr(True) is
            ' "Verdadero" in Spanish locales and "True" in English, and LCase
            ' normalises to a locale-specific form that generic parsers reject.
            ' Use the JSON spec literals explicitly.
            If value Then
                JsonValue = "true"
            Else
                JsonValue = "false"
            End If
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            JsonValue = Replace(CStr(value), ",", ".")
        Case Else
            JsonValue = """" & EscapeJsonString(CStr(value)) & """"
    End Select
End Function
