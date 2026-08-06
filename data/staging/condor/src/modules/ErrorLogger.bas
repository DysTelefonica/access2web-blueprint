Attribute VB_Name = "ErrorLogger"

Option Explicit
' ==========================================================================
' MÓDULO: ErrorLogger.bas
' RESPONSABILIDAD: Logging centralizado de errores a tbLogErrores.
'                 NUNCA lanza errores — todo wrapped en On Error Resume Next.
' ==========================================================================

' Log error to tbLogErrores — wrapped to NEVER throw
Public Sub LogToTable(ByRef condorErr As CondorError)
    On Error Resume Next  ' SILENT — intentionally
    
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim logErr As New LogError
    
    Set db = getdb()
    Set rcd = db.OpenRecordset("tbLogErrores", dbOpenDynaset)
    
    With logErr
        .fechaHora = Now()
        .usuario = IIf(m_ObjUsuarioActivo Is Nothing, "Sistema", m_ObjUsuarioActivo.nombre)
        .modulo = condorErr.source
        .procedimiento = IIf(InStrRev(condorErr.source, ".") > 0, Mid(condorErr.source, InStrRev(condorErr.source, ".") + 1), condorErr.source)
        .numeroError = condorErr.Number
        .descripcionError = condorErr.description
        .contexto = condorErr.CallStackToString()
    End With
    
    rcd.AddNew
    Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, logErr)
    rcd.Update
    
    rcd.Close
    Set rcd = Nothing
    Set db = Nothing
End Sub

' Log from global error object (convenience wrapper)
Public Sub LogFromGlobal()
    On Error Resume Next
    If g_objLastError Is Nothing Then Exit Sub
    Call LogToTable(g_objLastError)
End Sub

