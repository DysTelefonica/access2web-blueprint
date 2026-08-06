Attribute VB_Name = "Test_ConfigFixtures"
Option Compare Database
Option Explicit

Public Type TbConfiguracionSnapshot
    HasRecord As Boolean
    BackendActivo As String
    BackendProduccion As String
    BackendSandbox As String
    PasswordBackend As String
    IDAplicacion As String
    RutaDirectorioAplicacion_PROD As String
    RutaDirectorioAplicacion_LOCAL As String
    EnPruebas As String
End Type

Public Function SnapshotSingleConfigRow(ByRef p_Snapshot As TbConfiguracionSnapshot, _
                                        Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT TOP 2 * FROM TbConfiguracionBackends", dbOpenDynaset)

    If rs.EOF Then
        p_Snapshot.HasRecord = False
        SnapshotSingleConfigRow = True
        GoTo Cleanup
    End If

    p_Snapshot.HasRecord = True
    p_Snapshot.BackendActivo = CStr(Nz(rs.Fields("BackendActivo").value, ""))
    p_Snapshot.BackendProduccion = CStr(Nz(rs.Fields("BackendProduccion").value, ""))
    p_Snapshot.BackendSandbox = CStr(Nz(rs.Fields("BackendSandbox").value, ""))
    p_Snapshot.PasswordBackend = CStr(Nz(rs.Fields("PasswordBackend").value, ""))
    p_Snapshot.IDAplicacion = CStr(Nz(rs.Fields("IDAplicacion").value, ""))
    p_Snapshot.RutaDirectorioAplicacion_PROD = CStr(Nz(rs.Fields("RutaDirectorioAplicacion_PROD").value, ""))
    p_Snapshot.RutaDirectorioAplicacion_LOCAL = CStr(Nz(rs.Fields("RutaDirectorioAplicacion_LOCAL").value, ""))
    p_Snapshot.EnPruebas = CStr(Nz(rs.Fields("EnPruebas").value, ""))

    rs.MoveNext
    If Not rs.EOF Then
        p_Error = "TESTS BLOCKED: TbConfiguracionBackends tiene más de un registro; la suite no puede snapshotear el entorno de forma segura"
        GoTo Cleanup
    End If

    SnapshotSingleConfigRow = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

EH:
    p_Error = "SnapshotSingleConfigRow: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

Public Function ClearConfigRows(Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""
    CurrentDb.Execute "DELETE FROM TbConfiguracionBackends", dbFailOnError
    ClearConfigRows = True
    Exit Function
EH:
    p_Error = "ClearConfigRows: " & Err.Number & " - " & Err.Description
End Function

Public Function InsertConfigRow(ByVal p_BackendActivo As String, _
                                ByVal p_BackendProduccion As String, _
                                ByVal p_BackendSandbox As String, _
                                ByVal p_PasswordBackend As String, _
                                ByVal p_IDAplicacion As String, _
                                ByVal p_RutaDirectorioAplicacion_PROD As String, _
                                ByVal p_RutaDirectorioAplicacion_LOCAL As String, _
                                ByVal p_EnPruebas As String, _
                                Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("TbConfiguracionBackends", dbOpenDynaset)

    rs.AddNew
    rs.Fields("BackendActivo").value = UCase$(Trim$(p_BackendActivo))
    rs.Fields("BackendProduccion").value = Trim$(p_BackendProduccion)
    rs.Fields("BackendSandbox").value = Trim$(p_BackendSandbox)
    rs.Fields("PasswordBackend").value = p_PasswordBackend
    rs.Fields("IDAplicacion").value = Trim$(p_IDAplicacion)
    rs.Fields("RutaDirectorioAplicacion_PROD").value = NormalizarRuta(Trim$(p_RutaDirectorioAplicacion_PROD))
    rs.Fields("RutaDirectorioAplicacion_LOCAL").value = NormalizarRuta(Trim$(p_RutaDirectorioAplicacion_LOCAL))
    rs.Fields("EnPruebas").value = Trim$(p_EnPruebas)
    rs.Update

    InsertConfigRow = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

EH:
    p_Error = "InsertConfigRow: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

Public Function SeedSingleConfigRow(ByVal p_BackendActivo As String, _
                                    ByVal p_BackendProduccion As String, _
                                    ByVal p_BackendSandbox As String, _
                                    ByVal p_PasswordBackend As String, _
                                    ByVal p_IDAplicacion As String, _
                                    ByVal p_RutaDirectorioAplicacion_PROD As String, _
                                    ByVal p_RutaDirectorioAplicacion_LOCAL As String, _
                                    ByVal p_EnPruebas As String, _
                                    Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""

    If Not ClearConfigRows(p_Error) Then Exit Function
    If Not InsertConfigRow(p_BackendActivo, p_BackendProduccion, p_BackendSandbox, p_PasswordBackend, p_IDAplicacion, p_RutaDirectorioAplicacion_PROD, p_RutaDirectorioAplicacion_LOCAL, p_EnPruebas, p_Error) Then Exit Function

    SeedSingleConfigRow = True
    Exit Function
EH:
    p_Error = "SeedSingleConfigRow: " & Err.Number & " - " & Err.Description
End Function

Public Function SeedTwoConfigRows(Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""

    If Not ClearConfigRows(p_Error) Then Exit Function
    If Not InsertConfigRow("PROD", "\\fixture\prod\uno.accdb", "C:\fixture\local\uno.accdb", "", "12", "\\fixture\app\uno", "C:\fixture\app\uno", "No", p_Error) Then Exit Function
    If Not InsertConfigRow("LOCAL", "\\fixture\prod\dos.accdb", "C:\fixture\local\dos.accdb", "", "12", "\\fixture\app\dos", "C:\fixture\app\dos", "No", p_Error) Then Exit Function

    SeedTwoConfigRows = True
    Exit Function
EH:
    p_Error = "SeedTwoConfigRows: " & Err.Number & " - " & Err.Description
End Function

Public Function RestoreSingleConfigRow(ByRef p_Snapshot As TbConfiguracionSnapshot, _
                                       Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""

    If Not ClearConfigRows(p_Error) Then Exit Function
    If p_Snapshot.HasRecord Then
        If Not InsertConfigRow(p_Snapshot.BackendActivo, p_Snapshot.BackendProduccion, p_Snapshot.BackendSandbox, p_Snapshot.PasswordBackend, p_Snapshot.IDAplicacion, p_Snapshot.RutaDirectorioAplicacion_PROD, p_Snapshot.RutaDirectorioAplicacion_LOCAL, p_Snapshot.EnPruebas, p_Error) Then Exit Function
    End If

    Application.TempVars.RemoveAll
    ResetGlobals p_Error
    If p_Error <> "" Then Exit Function

    RestoreSingleConfigRow = True
    Exit Function
EH:
    p_Error = "RestoreSingleConfigRow: " & Err.Number & " - " & Err.Description
End Function

Public Function CountConfigRows(Optional ByRef p_Error As String = "") As Long
    On Error GoTo EH
    p_Error = ""
    CountConfigRows = CLng(Nz(DCount("*", "TbConfiguracionBackends"), 0))
    Exit Function
EH:
    p_Error = "CountConfigRows: " & Err.Number & " - " & Err.Description
End Function

