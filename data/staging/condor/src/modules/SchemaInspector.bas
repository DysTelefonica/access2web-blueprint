Attribute VB_Name = "SchemaInspector"
Option Compare Database
Option Explicit

' ============================================================================
' SchemaInspector — schema introspection utilities for test pre-checks
'
' Read-only helpers used to verify Access/DAO schema before writing tests
' that depend on FK constraints, unique indexes, or specific field types.
'
' Returns plain-text reports. Run via:
'   dysflow.dysflow_vba_execute({ procedureName: "Schema_Diag_DumpAll" })
' ============================================================================

Public Function Schema_Diag_DumpAll() As String
    On Error GoTo EH
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim rel As DAO.Relation
    Dim fieldLoop As DAO.Field
    Dim out As String

    Set db = CurrentDb
    out = "=== Schema_Diag_DumpAll ===" & vbCrLf

    For Each tdf In db.TableDefs
        If Left$(tdf.Name, 4) <> "MSys" And Left$(tdf.Name, 4) <> "~TMP" Then
            out = out & vbCrLf & "[TableDef] " & tdf.Name & vbCrLf
            For Each fld In tdf.Fields
                out = out & "  Field: " & fld.Name & " type=" & fld.Type & _
                    " size=" & fld.Size & " required=" & fld.Required & _
                    " allowZLS=" & fld.AllowZeroLength & vbCrLf
            Next fld
            For Each idx In tdf.Indexes
                Dim idxFields As String
                idxFields = ""
                For Each fieldLoop In idx.Fields
                    If Len(idxFields) > 0 Then idxFields = idxFields & ","
                    idxFields = idxFields & fieldLoop.Name
                Next fieldLoop
                out = out & "  Index: " & idx.Name & _
                    " primary=" & idx.Primary & _
                    " unique=" & idx.Unique & _
                    " required=" & idx.Required & _
                    " fields=[" & idxFields & "]" & vbCrLf
            Next idx
        End If
    Next tdf

    out = out & vbCrLf & "=== Relations ===" & vbCrLf
    For Each rel In db.Relations
        Dim relFields As String
        relFields = ""
        For Each fieldLoop In rel.Fields
            If Len(relFields) > 0 Then relFields = relFields & ","
            relFields = relFields & fieldLoop.Name & "->" & fieldLoop.ForeignName
        Next fieldLoop
        out = out & "  Relation: " & rel.Name & _
            " table=" & rel.Table & _
            " foreignTable=" & rel.ForeignTable & _
            " fields=[" & relFields & "]" & vbCrLf
    Next rel

    Schema_Diag_DumpAll = out
    Exit Function

EH:
    Schema_Diag_DumpAll = "Schema_Diag_DumpAll ERR " & Err.Number & " - " & Err.Description
End Function

Public Function Schema_Diag_AdjuntosOnly() As String
    On Error GoTo EH
    Dim dbFrontend As DAO.Database
    Dim dbBackend As DAO.Database
    Dim tdfBackend As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim fieldLoop As DAO.Field
    Dim out As String
    Dim rs As DAO.Recordset
    Dim backendPath As String
    Dim backendPwd As String

    Set dbFrontend = CurrentDb

    ' Read backend path from TbConfiguracionBackends
    Set rs = dbFrontend.OpenRecordset( _
        "SELECT TOP 1 BackendActivo, BackendSandbox, BackendTest, PasswordBackend " & _
        "FROM TbConfiguracionBackends WHERE Habilitado=True", dbOpenSnapshot)
    If rs.EOF Then
        out = "WARN: TbConfiguracionBackends has no enabled row"
        Schema_Diag_AdjuntosOnly = out
        Exit Function
    End If
    backendPwd = Nz(rs!PasswordBackend, "")
    rs.Close
    Set rs = Nothing

    ' Resolve backend path by trying each candidate in order. Skip symbolic values like "TEST"/"PROD".
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim candidates(1 To 3) As String
    Set rs = dbFrontend.OpenRecordset( _
        "SELECT TOP 1 BackendActivo, BackendSandbox, BackendTest " & _
        "FROM TbConfiguracionBackends WHERE Habilitado=True", dbOpenSnapshot)
    candidates(1) = Trim$(Nz(rs!BackendActivo, ""))
    candidates(2) = Trim$(Nz(rs!BackendSandbox, ""))
    candidates(3) = Trim$(Nz(rs!BackendTest, ""))
    rs.Close
    Set rs = Nothing
    Dim i As Integer
    backendPath = ""
    For i = 1 To 3
        If Len(candidates(i)) > 0 Then
            If fso.FileExists(candidates(i)) Then
                backendPath = candidates(i)
                Exit For
            End If
        End If
    Next i
    Set fso = Nothing
    out = "--- Backend path candidates ---" & vbCrLf
    out = out & "BackendActivo=[" & candidates(1) & "]" & vbCrLf
    out = out & "BackendSandbox=[" & candidates(2) & "]" & vbCrLf
    out = out & "BackendTest=[" & candidates(3) & "]" & vbCrLf
    out = out & "Resolved=" & backendPath & vbCrLf
    out = out & "PasswordSet=" & (Len(backendPwd) > 0) & vbCrLf
    If Len(backendPath) = 0 Then
        out = out & "ERROR: no candidate path resolves to an existing file" & vbCrLf
        Schema_Diag_AdjuntosOnly = out
        Exit Function
    End If

    ' Open backend directly (tbAdjuntos is backend-only, not a frontend linked table)
    If Len(backendPwd) > 0 Then
        Set dbBackend = DBEngine.Workspaces(0).OpenDatabase(backendPath, False, True, ";PWD=" & backendPwd)
    Else
        Set dbBackend = DBEngine.Workspaces(0).OpenDatabase(backendPath, False, True)
    End If

    out = out & vbCrLf & "=== tbAdjuntos in BACKEND ===" & vbCrLf
    Set tdfBackend = dbBackend.TableDefs("tbAdjuntos")
    For Each fld In tdfBackend.Fields
        out = out & "Field: " & fld.Name & " type=" & fld.Type & _
            " size=" & fld.Size & " required=" & fld.Required & _
            " allowZLS=" & fld.AllowZeroLength & vbCrLf
    Next fld
    out = out & "--- Indexes ---" & vbCrLf
    For Each idx In tdfBackend.Indexes
        Dim idxFieldsB As String
        idxFieldsB = ""
        For Each fieldLoop In idx.Fields
            If Len(idxFieldsB) > 0 Then idxFieldsB = idxFieldsB & ","
            idxFieldsB = idxFieldsB & fieldLoop.Name
        Next fieldLoop
        out = out & "Index: " & idx.Name & _
            " primary=" & idx.Primary & _
            " unique=" & idx.Unique & _
            " required=" & idx.Required & _
            " fields=[" & idxFieldsB & "]" & vbCrLf
    Next idx

    dbBackend.Close
    Set dbBackend = Nothing
    Schema_Diag_AdjuntosOnly = out
    Exit Function

EH:
    Schema_Diag_AdjuntosOnly = "Schema_Diag_AdjuntosOnly ERR " & Err.Number & " - " & Err.Description & _
        vbCrLf & "out-so-far=[" & out & "]" & vbCrLf & "backendPath=[" & backendPath & "]"
End Function

Public Function Schema_Diag_ExpedienteToSolicitudFK() As String
    On Error GoTo EH
    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim out As String

    Set db = CurrentDb
    out = "=== FK from TbExpedientes to tbSolicitudes ===" & vbCrLf
    For Each rel In db.Relations
        If (rel.Table = "TbExpedientes" And rel.ForeignTable = "tbSolicitudes") Or _
           (rel.Table = "tbSolicitudes" And rel.ForeignTable = "TbExpedientes") Then
            Dim fieldLoop As DAO.Field
            Dim relFields As String
            relFields = ""
            For Each fieldLoop In rel.Fields
                If Len(relFields) > 0 Then relFields = relFields & ","
                relFields = relFields & fieldLoop.Name & "->" & fieldLoop.ForeignName
            Next fieldLoop
            out = out & "FOUND: " & rel.Name & _
                " " & rel.Table & " -> " & rel.ForeignTable & _
                " fields=[" & relFields & "]" & vbCrLf
        End If
    Next rel
    If InStr(out, "FOUND") = 0 Then
        out = out & "NOT FOUND: no formal FK declared between TbExpedientes and tbSolicitudes via Relations" & vbCrLf
        out = out & "Implication: child table INSERTs will NOT enforce referential integrity via the formal FK; " & _
            "however, any explicit check in stored queries or application code may still fail at insert time." & vbCrLf
    End If

    Schema_Diag_ExpedienteToSolicitudFK = out
    Exit Function

EH:
    Schema_Diag_ExpedienteToSolicitudFK = "Schema_Diag_ExpedienteToSolicitudFK ERR " & Err.Number & " - " & Err.Description
End Function
