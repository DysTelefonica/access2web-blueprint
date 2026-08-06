Attribute VB_Name = "PriorizacionRepository"
Option Compare Database
Option Explicit

' Issue #36 repository seam for frontend/local TbAuxPriorizacion.
Public Function Priorizacion_ClearTemp(Optional ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim db As DAO.Database

    p_Error = vbNullString
    Set db = CurrentDb
    db.Execute "DELETE FROM TbAuxPriorizacion", dbFailOnError

    Priorizacion_ClearTemp = True
    Set db = Nothing
    Exit Function

EH:
    If p_Error = "" Then p_Error = "Priorizacion_ClearTemp: " & Err.Description
    Priorizacion_ClearTemp = False
End Function

Public Function Priorizacion_InsertSnapshotRow( _
    ByVal p_IDRiesgo As Long, _
    ByVal p_Pri As String, _
    ByVal p_PriEdAnterior As String, _
    ByVal p_Codigo As String, _
    ByVal p_Descripcion As String, _
    Optional ByVal p_CausaRaiz As String = "", _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo EH

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    p_Error = vbNullString
    Set db = CurrentDb
    Set rs = db.OpenRecordset("TbAuxPriorizacion", dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    rs.Fields("IDRiesgo").Value = p_IDRiesgo
    rs.Fields("Pri").Value = p_Pri
    rs.Fields("PriEdAnterior").Value = p_PriEdAnterior
    rs.Fields("Codigo").Value = p_Codigo
    rs.Fields("Descripcion").Value = p_Descripcion
    rs.Fields("CausaRaiz").Value = p_CausaRaiz
    rs.Update

    Priorizacion_InsertSnapshotRow = True

SALIR:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function

EH:
    If p_Error = "" Then p_Error = "Priorizacion_InsertSnapshotRow: " & Err.Description
    Priorizacion_InsertSnapshotRow = False
    Resume SALIR
End Function

Public Function Priorizacion_ReadTempPriorities(Optional ByRef p_Error As String) As Scripting.Dictionary
    On Error GoTo EH

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim priorities As Scripting.Dictionary
    Dim riskId As Variant

    p_Error = vbNullString
    Set priorities = New Scripting.Dictionary
    priorities.CompareMode = TextCompare

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT IDRiesgo, Pri FROM TbAuxPriorizacion WHERE IDRiesgo IS NOT NULL", _
        dbOpenSnapshot)

    Do While Not rs.EOF
        riskId = rs.Fields("IDRiesgo").Value
        priorities(CStr(CLng(riskId))) = Nz(rs.Fields("Pri").Value, vbNullString)
        rs.MoveNext
    Loop

    Set Priorizacion_ReadTempPriorities = priorities

SALIR:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set priorities = Nothing
    Set db = Nothing
    Exit Function

EH:
    If p_Error = "" Then p_Error = "Priorizacion_ReadTempPriorities: " & Err.Description
    Set Priorizacion_ReadTempPriorities = Nothing
    Resume SALIR
End Function

Public Function Priorizacion_UpdateTempPriority( _
    ByVal p_IDRiesgo As Long, _
    ByVal p_Pri As String, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo EH

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    p_Error = vbNullString
    Set db = CurrentDb
    Set qdf = db.CreateQueryDef(vbNullString, _
        "PARAMETERS p_Pri TEXT(255), p_IDRiesgo LONG; " & _
        "UPDATE TbAuxPriorizacion SET Pri = [p_Pri] WHERE IDRiesgo = [p_IDRiesgo]")
    qdf.Parameters("p_Pri").Value = p_Pri
    qdf.Parameters("p_IDRiesgo").Value = p_IDRiesgo
    qdf.Execute dbFailOnError

    Priorizacion_UpdateTempPriority = True

SALIR:
    Set qdf = Nothing
    Set db = Nothing
    Exit Function

EH:
    If p_Error = "" Then p_Error = "Priorizacion_UpdateTempPriority: " & Err.Description
    Priorizacion_UpdateTempPriority = False
    Resume SALIR
End Function
