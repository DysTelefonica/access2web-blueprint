Attribute VB_Name = "E2EFamilyExportHelper"
Option Compare Database
Option Explicit

Private Const E2E_FAMILY_ORDINAL_MAX_NULL_SORT As Long = 2147483647

Private m_ForceOrdinalPreparationFailureForTest As Boolean

Public Sub E2EFamilyExport_SetOrdinalPreparationFailureForTest(ByVal p_Force As Boolean)
    m_ForceOrdinalPreparationFailureForTest = p_Force
End Sub

Public Function E2EFamilyExport_ExpandCsvToFamilyCsv( _
    ByVal p_Db As DAO.Database, _
    ByVal p_IDExpedientesCsv As String, _
    Optional ByRef p_Error As String) As String

    On Error GoTo ErrorHandler

    Dim selectedIds As Collection
    Dim outputIds As Object
    Dim i As Long
    Dim selectedId As Long
    Dim rootId As Long

    p_Error = ""
    If p_Db Is Nothing Then
        p_Error = "E2EFamilyExport_ExpandCsvToFamilyCsv: database is required"
        Exit Function
    End If

    Set selectedIds = ParseCsvIds(p_IDExpedientesCsv)
    If selectedIds.Count = 0 Then
        p_Error = "E2EFamilyExport_ExpandCsvToFamilyCsv: lista de IDs vacía o inválida"
        Exit Function
    End If

    Set outputIds = CreateObject("Scripting.Dictionary")
    For i = 1 To selectedIds.Count
        selectedId = CLng(selectedIds.Item(i))
        rootId = FindRootExpedienteId(p_Db, selectedId, p_Error)
        If p_Error <> "" Then Exit Function
        If rootId > 0 Then ExpandDescendantsFromRoot p_Db, rootId, outputIds, p_Error
        If p_Error <> "" Then Exit Function
    Next i

    If outputIds.Count = 0 Then
        p_Error = "E2EFamilyExport_ExpandCsvToFamilyCsv: ningún ID corresponde a un expediente existente"
        Exit Function
    End If

    E2EFamilyExport_ExpandCsvToFamilyCsv = DictionaryKeysToCsv(outputIds)
    Exit Function

ErrorHandler:
    p_Error = "E2EFamilyExport_ExpandCsvToFamilyCsv: " & Err.Description
    E2EFamilyExport_ExpandCsvToFamilyCsv = ""
End Function

Public Function E2EFamilyExport_AssignMissingOrdinals( _
    ByVal p_Db As DAO.Database, _
    ByVal p_IDExpedientesCsv As String, _
    Optional ByRef p_Error As String) As Boolean

    On Error GoTo ErrorHandler

    Dim ids As Collection
    Dim whereClause As String
    Dim nextOrdinal As Long
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef

    p_Error = ""
    E2EFamilyExport_AssignMissingOrdinals = False

    If p_Db Is Nothing Then
        p_Error = "E2EFamilyExport_AssignMissingOrdinals: database is required"
        Exit Function
    End If

    Set ids = ParseCsvIds(p_IDExpedientesCsv)
    If ids.Count = 0 Then
        p_Error = "E2EFamilyExport_AssignMissingOrdinals: lista de IDs vacía o inválida"
        Exit Function
    End If

    nextOrdinal = GetMaxOrdinalE2E(p_Db) + 1
    whereClause = BuildParameterizedIdPredicate(ids, "IDExpediente")
    Set qdf = p_Db.CreateQueryDef("", _
        "PARAMETERS " & BuildIdParameterDeclarations(ids) & "; " & _
        "SELECT IDExpediente FROM TbExpedientes " & _
        "WHERE (" & whereClause & ") AND OrdinalE2E IS NULL " & _
        "ORDER BY IDExpediente;")
    ApplyIdParameters qdf, ids
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)

    Set qdf = Nothing

    Set qdf = p_Db.CreateQueryDef("", _
        "PARAMETERS p_Ordinal Long, p_IDExpediente Long; " & _
        "UPDATE TbExpedientes SET OrdinalE2E=[p_Ordinal] WHERE IDExpediente=[p_IDExpediente];")

    Do While Not rs.EOF
        qdf.Parameters("p_Ordinal").Value = nextOrdinal
        qdf.Parameters("p_IDExpediente").Value = CLng(rs.Fields("IDExpediente").Value)
        qdf.Execute dbFailOnError

        If m_ForceOrdinalPreparationFailureForTest Then
            p_Error = "E2EFamilyExport_AssignMissingOrdinals: forced ordinal preparation failure"
            Err.Raise vbObjectError + 1001, "E2EFamilyExportHelper", p_Error
        End If

        nextOrdinal = nextOrdinal + 1
        rs.MoveNext
    Loop

    E2EFamilyExport_AssignMissingOrdinals = True

SafeExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    Exit Function

ErrorHandler:
    If p_Error = "" Then p_Error = "E2EFamilyExport_AssignMissingOrdinals: " & Err.Description
    E2EFamilyExport_AssignMissingOrdinals = False
    Resume SafeExit
End Function

Private Function ParseCsvIds(ByVal p_Csv As String) As Collection
    Dim ids As New Collection
    Dim seen As Object
    Dim parts() As String
    Dim i As Long
    Dim token As String
    Dim idValue As Long
    Dim normalized As String

    Set seen = CreateObject("Scripting.Dictionary")
    normalized = Trim$(Replace(Replace(p_Csv, vbCr, ""), vbLf, ""))
    If Len(normalized) > 0 Then
        parts = Split(normalized, ",")
        For i = LBound(parts) To UBound(parts)
            token = Trim$(parts(i))
            If Len(token) > 0 Then
                If IsNumeric(token) Then
                    idValue = CLng(token)
                    If idValue > 0 Then
                        If Not seen.Exists(CStr(idValue)) Then
                            seen.Add CStr(idValue), True
                            ids.Add idValue
                        End If
                    End If
                End If
            End If
        Next i
    End If

    Set ParseCsvIds = ids
End Function

Private Function FindRootExpedienteId(ByVal p_Db As DAO.Database, ByVal p_IDExpediente As Long, ByRef p_Error As String) As Long
    Dim ancestorPath As Object
    Dim currentId As Long
    Dim parentId As Variant

    Set ancestorPath = CreateObject("Scripting.Dictionary")
    currentId = p_IDExpediente

    Do While currentId > 0
        If ancestorPath.Exists(CStr(currentId)) Then
            FindRootExpedienteId = currentId
            Exit Function
        End If

        ancestorPath.Add CStr(currentId), True
        parentId = GetParentExpedienteId(p_Db, currentId, p_Error)
        If p_Error <> "" Then Exit Function
        If IsEmpty(parentId) Then Exit Function
        If IsNull(parentId) Then
            FindRootExpedienteId = currentId
            Exit Function
        End If

        currentId = CLng(parentId)
    Loop
End Function

Private Function GetParentExpedienteId(ByVal p_Db As DAO.Database, ByVal p_IDExpediente As Long, ByRef p_Error As String) As Variant
    On Error GoTo ErrorHandler

    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset

    Set qdf = p_Db.CreateQueryDef("", _
        "PARAMETERS p_IDExpediente Long; " & _
        "SELECT IDExpedientePadre FROM TbExpedientes WHERE IDExpediente=[p_IDExpediente];")
    qdf.Parameters("p_IDExpediente").Value = p_IDExpediente
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)

    If rs.EOF Then
        GetParentExpedienteId = Empty
    Else
        GetParentExpedienteId = rs.Fields("IDExpedientePadre").Value
    End If

SafeExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    Exit Function

ErrorHandler:
    p_Error = "GetParentExpedienteId: " & Err.Description
    GetParentExpedienteId = Empty
    Resume SafeExit
End Function

Private Sub ExpandDescendantsFromRoot( _
    ByVal p_Db As DAO.Database, _
    ByVal p_RootId As Long, _
    ByVal p_OutputIds As Object, _
    ByRef p_Error As String)

    Dim descendantQueued As Object
    Dim descendantProcessed As Object
    Dim queue As Collection
    Dim index As Long
    Dim currentId As Long
    Dim children As Collection
    Dim i As Long
    Dim childId As Long

    Set descendantQueued = CreateObject("Scripting.Dictionary")
    Set descendantProcessed = CreateObject("Scripting.Dictionary")
    Set queue = New Collection

    AddIdToDictionary p_OutputIds, p_RootId
    AddIdToDictionary descendantQueued, p_RootId
    queue.Add p_RootId

    index = 1
    Do While index <= queue.Count
        currentId = CLng(queue.Item(index))
        index = index + 1

        If Not descendantProcessed.Exists(CStr(currentId)) Then
            descendantProcessed.Add CStr(currentId), True
            Set children = GetChildExpedienteIds(p_Db, currentId, p_Error)
            If p_Error <> "" Then Exit Sub

            For i = 1 To children.Count
                childId = CLng(children.Item(i))
                AddIdToDictionary p_OutputIds, childId
                If Not descendantQueued.Exists(CStr(childId)) Then
                    descendantQueued.Add CStr(childId), True
                    queue.Add childId
                End If
            Next i
        End If
    Loop
End Sub

Private Function GetChildExpedienteIds(ByVal p_Db As DAO.Database, ByVal p_IDPadre As Long, ByRef p_Error As String) As Collection
    On Error GoTo ErrorHandler

    Dim children As New Collection
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset

    Set qdf = p_Db.CreateQueryDef("", _
        "PARAMETERS p_IDPadre Long; " & _
        "SELECT IDExpediente FROM TbExpedientes " & _
        "WHERE IDExpedientePadre=[p_IDPadre] " & _
        "ORDER BY Nz(OrdinalE2E, " & CStr(E2E_FAMILY_ORDINAL_MAX_NULL_SORT) & "), IDExpediente;")
    qdf.Parameters("p_IDPadre").Value = p_IDPadre
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)

    Do While Not rs.EOF
        children.Add CLng(rs.Fields("IDExpediente").Value)
        rs.MoveNext
    Loop

SafeExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set GetChildExpedienteIds = children
    Exit Function

ErrorHandler:
    p_Error = "GetChildExpedienteIds: " & Err.Description
    Resume SafeExit
End Function

Private Function GetMaxOrdinalE2E(ByVal p_Db As DAO.Database) As Long
    Dim rs As DAO.Recordset

    Set rs = p_Db.OpenRecordset("SELECT Max(OrdinalE2E) AS MaxOrdinal FROM TbExpedientes WHERE OrdinalE2E IS NOT NULL;", dbOpenSnapshot)
    If Not rs.EOF Then GetMaxOrdinalE2E = CLng(Nz(rs.Fields("MaxOrdinal").Value, 0))
    rs.Close
    Set rs = Nothing
End Function

Private Function BuildIdParameterDeclarations(ByVal p_Ids As Collection) As String
    Dim i As Long

    For i = 1 To p_Ids.Count
        If BuildIdParameterDeclarations <> "" Then BuildIdParameterDeclarations = BuildIdParameterDeclarations & ", "
        BuildIdParameterDeclarations = BuildIdParameterDeclarations & "p_ID" & CStr(i) & " Long"
    Next i
End Function

Private Function BuildParameterizedIdPredicate(ByVal p_Ids As Collection, ByVal p_FieldName As String) As String
    Dim i As Long

    For i = 1 To p_Ids.Count
        If BuildParameterizedIdPredicate <> "" Then BuildParameterizedIdPredicate = BuildParameterizedIdPredicate & " OR "
        BuildParameterizedIdPredicate = BuildParameterizedIdPredicate & p_FieldName & "=[p_ID" & CStr(i) & "]"
    Next i
End Function

Private Sub ApplyIdParameters(ByVal p_QueryDef As DAO.QueryDef, ByVal p_Ids As Collection)
    Dim i As Long

    For i = 1 To p_Ids.Count
        p_QueryDef.Parameters("p_ID" & CStr(i)).Value = CLng(p_Ids.Item(i))
    Next i
End Sub

Private Sub AddIdToDictionary(ByVal p_Dictionary As Object, ByVal p_IDExpediente As Long)
    If Not p_Dictionary.Exists(CStr(p_IDExpediente)) Then p_Dictionary.Add CStr(p_IDExpediente), p_IDExpediente
End Sub

Private Function DictionaryKeysToCsv(ByVal p_Dictionary As Object) As String
    Dim key As Variant

    For Each key In p_Dictionary.Keys
        If DictionaryKeysToCsv <> "" Then DictionaryKeysToCsv = DictionaryKeysToCsv & ","
        DictionaryKeysToCsv = DictionaryKeysToCsv & CStr(p_Dictionary.Item(key))
    Next key
End Function
