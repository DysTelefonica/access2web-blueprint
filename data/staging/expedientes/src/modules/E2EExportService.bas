Attribute VB_Name = "E2EExportService"
Option Compare Database
Option Explicit

Public Function EnsureE2EExportTraceabilitySchema(Optional ByRef p_Error As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    p_Error = ""
    EnsureE2EExportTraceabilitySchema = False

    If Not EnsureE2EBatchManagementSchema(p_Error) Then Exit Function

    Set db = getdb(p_Error)
    If p_Error <> "" Then Exit Function
    If db Is Nothing Then
        p_Error = "EnsureE2EExportTraceabilitySchema: getdb returned Nothing"
        Exit Function
    End If

    If Not RequireTraceabilityTextField(db, "HashActual", p_Error) Then Exit Function
    If Not RequireTraceabilityTextField(db, "HashUltimaExportacion", p_Error) Then Exit Function

    EnsureE2EExportTraceabilitySchema = True
    Exit Function

ErrorHandler:
    p_Error = "EnsureE2EExportTraceabilitySchema: " & Err.Description
    EnsureE2EExportTraceabilitySchema = False
End Function

Public Function MigrateE2EExportTraceabilitySchema(Optional ByRef p_Error As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    p_Error = ""
    MigrateE2EExportTraceabilitySchema = False

    If Not EnsureE2EBatchManagementSchema(p_Error) Then Exit Function


    Set db = getdb(p_Error)
    If p_Error <> "" Then Exit Function
    If db Is Nothing Then
        p_Error = "MigrateE2EExportTraceabilitySchema: getdb returned Nothing"
        Exit Function
    End If

    If Not EnsureTraceabilityTextField(db, "HashActual", p_Error) Then Exit Function
    If Not EnsureTraceabilityTextField(db, "HashUltimaExportacion", p_Error) Then Exit Function

    MigrateE2EExportTraceabilitySchema = True
    Exit Function

ErrorHandler:
    p_Error = "MigrateE2EExportTraceabilitySchema: " & Err.Description
    MigrateE2EExportTraceabilitySchema = False
End Function

Public Function PreviewAutomaticE2EExportSelectionCsv( _
    Optional ByRef p_Error As String, _
    Optional ByVal p_CandidateIdsCsv As String = "") As String

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim result As String
    Dim ids As Collection
    Dim trimmedCsv As String
    Dim idValue As Long

    On Error GoTo ErrorHandler
    p_Error = ""

    If Not EnsureE2EExportTraceabilitySchema(p_Error) Then Exit Function
    Set db = getdb(p_Error)
    If p_Error <> "" Then Exit Function

    trimmedCsv = Trim$(Replace(Replace(p_CandidateIdsCsv, vbCr, ""), vbLf, ""))
    Set ids = New Collection
    If Len(trimmedCsv) > 0 Then
        Dim parts() As String
        Dim i As Long
        Dim token As String
        parts = Split(trimmedCsv, ",")
        For i = LBound(parts) To UBound(parts)
            token = Trim$(parts(i))
            If Len(token) > 0 Then
                If Not IsNumeric(token) Then
                    p_Error = "PreviewAutomaticE2EExportSelectionCsv: candidate id is not numeric"
                    Exit Function
                End If
                idValue = CLng(token)
                If idValue > 0 Then ids.Add idValue
            End If
        Next i
    End If

    If ids.Count = 0 Then
        ' Empty CSV path: preserve original "all candidates" semantic (no candidate predicate).
        Set rs = db.OpenRecordset( _
            "SELECT IDExpediente FROM TbExpedientes " & _
            "WHERE OrdinalE2E IS NOT NULL " & _
            "AND (Nz(HashUltimaExportacion,'')='' OR Nz(HashActual,'')<>Nz(HashUltimaExportacion,'')) " & _
            "ORDER BY IDExpediente;", dbOpenSnapshot)
    Else
        ' Non-empty CSV path: parameterized DAO QueryDef with OR-chain predicate.
        Set qdf = db.CreateQueryDef("", _
            "PARAMETERS " & BuildCandidateIdParameterDeclarations(ids) & "; " & _
            "SELECT IDExpediente FROM TbExpedientes " & _
            "WHERE OrdinalE2E IS NOT NULL " & _
            "AND (" & BuildCandidateIdPredicate(ids, "IDExpediente") & ") " & _
            "AND (Nz(HashUltimaExportacion,'')='' OR Nz(HashActual,'')<>Nz(HashUltimaExportacion,'')) " & _
            "ORDER BY IDExpediente;")
        Dim k As Long
        For k = 1 To ids.Count
            qdf.Parameters("p_ID" & CStr(k)).Value = CLng(ids.Item(k))
        Next k
        Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    End If

    Do While Not rs.EOF
        AppendCsv result, CStr(rs.Fields("IDExpediente").value)
        rs.MoveNext
    Loop
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    On Error GoTo 0

    PreviewAutomaticE2EExportSelectionCsv = result
    Exit Function

ErrorHandler:
    p_Error = "PreviewAutomaticE2EExportSelectionCsv: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdf = Nothing
End Function

Public Function PreviewManualE2EExportSelectionCsv( _
    ByVal p_UsuarioConectado As String, _
    ByVal p_SessionId As String, _
    Optional ByRef p_Error As String) As String

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim result As String

    On Error GoTo ErrorHandler
    p_Error = ""

    If Not EnsureE2EExportTraceabilitySchema(p_Error) Then Exit Function
    Set db = getdb(p_Error)
    If p_Error <> "" Then Exit Function

    Set rs = db.OpenRecordset( _
        "SELECT IDExpediente FROM TbE2EExportSeleccionTemp " & _
        "WHERE UsuarioConectado='" & SqlStr(p_UsuarioConectado) & "' " & _
        "AND SessionId='" & SqlStr(p_SessionId) & "' " & _
        "ORDER BY IDExpediente;", dbOpenSnapshot)

    Do While Not rs.EOF
        AppendCsv result, CStr(rs.Fields("IDExpediente").value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    PreviewManualE2EExportSelectionCsv = result
    Exit Function

ErrorHandler:
    p_Error = "PreviewManualE2EExportSelectionCsv: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Public Function GenerateE2EExportJsonForBatch( _
    ByVal p_IDBatch As Long, _
    ByVal p_IdsCsv As String, _
    Optional ByRef p_Error As String) As String

    On Error GoTo ErrorHandler

    p_Error = ""
    GenerateE2EExportJsonForBatch = GenerarJsonE2ECanonicoPorListaConFamiliaParaBatch( _
        CLng(p_IDBatch), CStr(p_IdsCsv), p_Error)
    Exit Function

ErrorHandler:
    p_Error = "GenerateE2EExportJsonForBatch: " & Err.Description
    GenerateE2EExportJsonForBatch = ""
End Function

Public Function ExportE2EAutomatic( _
    ByVal p_UsuarioConectado As String, _
    ByVal p_SessionId As String, _
    Optional ByVal p_DestinationPath As String = "", _
    Optional ByVal p_FileWriterMode As String = "", _
    Optional ByRef p_Error As String, _
    Optional ByVal p_CandidateIdsCsv As String = "") As String

    Dim db As DAO.Database
    Dim idsCsv As String
    Dim idBatch As Long
    Dim jsonText As String
    Dim ok As Boolean
    Dim filePath As String
    Dim fileError As String

    On Error GoTo ErrorHandler
    p_Error = ""

    If Not EnsureE2EExportTraceabilitySchema(p_Error) Then Exit Function
    Set db = getdb(p_Error)
    If p_Error <> "" Then Exit Function

    idsCsv = PreviewAutomaticE2EExportSelectionCsv(p_Error, p_CandidateIdsCsv)
    If p_Error <> "" Then Exit Function


    idBatch = CreateExportBatch(db, p_UsuarioConectado, p_SessionId, idsCsv)
    jsonText = GenerateE2EExportJsonForBatch(idBatch, idsCsv, p_Error)
    If p_Error <> "" Then Exit Function

    If LCase$(Trim$(p_FileWriterMode)) = "force-fail" Then
        fileError = "forced file failure"
        ok = False
    Else
        ok = WriteE2EExportFile(p_DestinationPath, idBatch, jsonText, filePath, fileError)
    End If

    If ok Then
        MarkExportSuccess db, idBatch, idsCsv
        ExportE2EAutomatic = "{""ok"":true,""IDBatch"":" & idBatch & ",""selectedIds"":""" & JsonEscape(idsCsv) & """,""outputPath"":""" & JsonEscape(filePath) & """,""json"":""" & JsonEscape(jsonText) & """,""error"":null}"
    Else
        db.Execute "UPDATE TbE2EExportBatch SET Estado='Error', ErrorMessage='" & SqlStr(fileError) & "' WHERE IDBatch=" & idBatch & ";", dbFailOnError
        ExportE2EAutomatic = "{""ok"":false,""IDBatch"":" & idBatch & ",""selectedIds"":""" & JsonEscape(idsCsv) & """,""outputPath"":""" & JsonEscape(filePath) & """,""json"":""" & JsonEscape(jsonText) & """,""error"":""" & JsonEscape(fileError) & """}"
    End If
    Exit Function

ErrorHandler:
    p_Error = "ExportE2EAutomatic: " & Err.Description
    ExportE2EAutomatic = "{""ok"":false,""IDBatch"":0,""selectedIds"":"""",""json"":"""",""error"":""" & JsonEscape(p_Error) & """}"
End Function

Private Function WriteE2EExportFile( _
    ByVal p_DestinationPath As String, _
    ByVal p_IDBatch As Long, _
    ByVal p_JsonText As String, _
    ByRef p_OutputPath As String, _
    ByRef p_Error As String) As Boolean

    Dim fso As Object

    On Error GoTo ErrorHandler

    p_Error = ""
    p_OutputPath = ""
    Set fso = CreateObject("Scripting.FileSystemObject")

    p_OutputPath = ResolveE2EExportOutputPath(fso, p_DestinationPath, p_IDBatch, p_Error)
    If p_Error <> "" Then Exit Function

    WriteE2EExportFile = WriteJsonUtf8File(p_OutputPath, p_JsonText, p_Error)
    Exit Function

ErrorHandler:
    p_Error = "WriteE2EExportFile: " & Err.Description
    On Error Resume Next
    WriteE2EExportFile = False
End Function

' ------------------------------------------------------------------------------------------
' Writes p_Text to p_FilePath as UTF-8 (with BOM-less ADODB.Stream), unescaping any \uXXXX
' sequences first. Used by WriteE2EExportFile and any other writer that needs literal
' non-ASCII chars on disk (the E2E consumer rejects ANSI/Win-1252 + literal \uXXXX).
' Returns True on success; on failure sets p_Error and returns False.
' ------------------------------------------------------------------------------------------
Public Function WriteJsonUtf8File( _
    ByVal p_FilePath As String, _
    ByVal p_Text As String, _
    ByRef p_Error As String) As Boolean

    Dim stream As Object
    Dim unescapedContent As String

    On Error GoTo ErrorHandler

    p_Error = ""
    WriteJsonUtf8File = False

    unescapedContent = UnescapeUnicode(p_Text)

    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.WriteText unescapedContent
    stream.SaveToFile p_FilePath, 2 ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing

    WriteJsonUtf8File = True
    Exit Function

ErrorHandler:
    p_Error = "WriteJsonUtf8File: " & Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    WriteJsonUtf8File = False
End Function

Private Function ResolveE2EExportOutputPath( _
    ByVal p_Fso As Object, _
    ByVal p_DestinationPath As String, _
    ByVal p_IDBatch As Long, _
    ByRef p_Error As String) As String

    Dim destination As String
    Dim parentFolder As String

    destination = Trim$(p_DestinationPath)
    If destination = "" Then
        p_Error = "destination path is required"
        Exit Function
    End If

    If p_Fso.FolderExists(destination) Then
        ResolveE2EExportOutputPath = p_Fso.BuildPath(destination, "e2e_export_" & CStr(p_IDBatch) & ".json")
        Exit Function
    End If

    If Right$(destination, 1) = "\" Or Right$(destination, 1) = "/" Then
        p_Error = "destination folder does not exist"
        Exit Function
    End If

    parentFolder = p_Fso.GetParentFolderName(destination)
    If parentFolder = "" Or Not p_Fso.FolderExists(parentFolder) Then
        p_Error = "destination folder does not exist"
        Exit Function
    End If

    ResolveE2EExportOutputPath = destination
End Function

Private Function CreateExportBatch( _
    ByVal p_Db As DAO.Database, _
    ByVal p_UsuarioConectado As String, _
    ByVal p_SessionId As String, _
    ByVal p_IdsCsv As String) As Long

    Dim rs As DAO.Recordset
    Dim total As Long

    total = CsvCount(p_IdsCsv)
    p_Db.Execute "INSERT INTO TbE2EExportBatch (SessionId, UsuarioConectado, Estado, CreatedAt, StartedAt, TotalSeleccionados, TotalExportados) VALUES (" & _
                 "'" & SqlStr(p_SessionId) & "', '" & SqlStr(p_UsuarioConectado) & "', 'Creado', Now(), Now(), " & total & ", 0);", dbFailOnError
    Set rs = p_Db.OpenRecordset("SELECT @@IDENTITY AS NewId", dbOpenSnapshot)
    If Not rs.EOF Then CreateExportBatch = CLng(Nz(rs.Fields("NewId").value, 0))
    rs.Close
    Set rs = Nothing
End Function

Private Sub MarkExportSuccess(ByVal p_Db As DAO.Database, ByVal p_IDBatch As Long, ByVal p_IdsCsv As String)
    Dim ids() As String
    Dim i As Long
    Dim idExp As Long
    Dim hashActual As String

    If Trim$(p_IdsCsv) = "" Then Exit Sub
    ids = Split(p_IdsCsv, ",")
    For i = LBound(ids) To UBound(ids)
        If Trim$(ids(i)) <> "" Then
            idExp = CLng(Trim$(ids(i)))
            hashActual = ReadHashActual(p_Db, idExp)
            p_Db.Execute "INSERT INTO TbE2EExportBatchDetalle (IDBatch, IDExpediente, OrdinalSeleccion, HashExportado, Estado, CreatedAt, ExportedAt) VALUES (" & _
                         p_IDBatch & ", " & idExp & ", " & (i + 1) & ", '" & SqlStr(hashActual) & "', 'Exitoso', Now(), Now());", dbFailOnError
            p_Db.Execute "UPDATE TbExpedientes SET HashUltimaExportacion='" & SqlStr(hashActual) & "' WHERE IDExpediente=" & idExp & ";", dbFailOnError
        End If
    Next i
    p_Db.Execute "UPDATE TbE2EExportBatch SET Estado='Exitoso', CompletedAt=Now(), TotalExportados=" & CsvCount(p_IdsCsv) & " WHERE IDBatch=" & p_IDBatch & ";", dbFailOnError
End Sub

Private Function ReadHashActual(ByVal p_Db As DAO.Database, ByVal p_IDExpediente As Long) As String
    Dim rs As DAO.Recordset
    Set rs = p_Db.OpenRecordset("SELECT HashActual FROM TbExpedientes WHERE IDExpediente=" & p_IDExpediente & ";", dbOpenSnapshot)
    If Not rs.EOF Then ReadHashActual = Nz(rs.Fields("HashActual").value, "")
    rs.Close
    Set rs = Nothing
End Function

Private Sub AppendCsv(ByRef p_Csv As String, ByVal p_Value As String)
    If p_Csv <> "" Then p_Csv = p_Csv & ","
    p_Csv = p_Csv & p_Value
End Sub

Private Function CsvCount(ByVal p_Csv As String) As Long
    If Trim$(p_Csv) = "" Then Exit Function
    CsvCount = UBound(Split(p_Csv, ",")) + 1
End Function

Private Function BuildCandidateIdParameterDeclarations(ByVal p_Ids As Collection) As String
    Dim i As Long

    For i = 1 To p_Ids.Count
        If BuildCandidateIdParameterDeclarations <> "" Then BuildCandidateIdParameterDeclarations = BuildCandidateIdParameterDeclarations & ", "
        BuildCandidateIdParameterDeclarations = BuildCandidateIdParameterDeclarations & "p_ID" & CStr(i) & " Long"
    Next i
End Function

Private Function BuildCandidateIdPredicate(ByVal p_Ids As Collection, ByVal p_FieldName As String) As String
    Dim i As Long

    For i = 1 To p_Ids.Count
        If BuildCandidateIdPredicate <> "" Then BuildCandidateIdPredicate = BuildCandidateIdPredicate & " OR "
        BuildCandidateIdPredicate = BuildCandidateIdPredicate & p_FieldName & "=[p_ID" & CStr(i) & "]"
    Next i
End Function

Private Function SqlStr(ByVal p_Value As String) As String
    SqlStr = Replace(Nz(p_Value, ""), "'", "''")
End Function

Private Function JsonEscape(ByVal p_Value As String) As String
    JsonEscape = Replace(Replace(Nz(p_Value, ""), "\", "\\"), Chr$(34), "\" & Chr$(34))
End Function

Private Function EnsureTraceabilityTextField( _
    ByVal p_Db As DAO.Database, _
    ByVal p_FieldName As String, _
    ByRef p_Error As String) As Boolean

    On Error GoTo ErrorHandler

    p_Error = ""
    If TraceabilityFieldExists(p_Db, p_FieldName) Then
        EnsureTraceabilityTextField = True
        Exit Function
    End If

    p_Db.Execute "ALTER TABLE TbExpedientes ADD COLUMN " & p_FieldName & " TEXT(64);", dbFailOnError
    EnsureTraceabilityTextField = True
    Exit Function

ErrorHandler:
    p_Error = "EnsureTraceabilityTextField(" & p_FieldName & "): " & Err.Description
    EnsureTraceabilityTextField = False
End Function

Private Function RequireTraceabilityTextField( _
    ByVal p_Db As DAO.Database, _
    ByVal p_FieldName As String, _
    ByRef p_Error As String) As Boolean

    p_Error = ""
    If TraceabilityFieldExists(p_Db, p_FieldName) Then
        RequireTraceabilityTextField = True
    Else
        p_Error = "TbExpedientes." & p_FieldName & " missing. Run MigrateE2EExportTraceabilitySchema explicitly before using E2E export."
        RequireTraceabilityTextField = False
    End If
End Function

Private Function TraceabilityFieldExists(ByVal p_Db As DAO.Database, ByVal p_FieldName As String) As Boolean
    On Error GoTo NotFound

    Dim fld As DAO.Field
    Set fld = p_Db.TableDefs("TbExpedientes").Fields(p_FieldName)
    TraceabilityFieldExists = True
    Exit Function

NotFound:
    TraceabilityFieldExists = False
End Function




