VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SnapshotServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' CLASE: SnapshotServicio
' DESCRIPCIÓN: Genera un snapshot JSON consolidado de una solicitud para
'              detectar cambios en los datos (Spec-105).
' ==========================================================================

Public Function GenerarSnapshotGlobal(ByVal idSolicitud As Long) As String
    ' Genera un JSON con los datos de todas las pestañas relevantes.
    ' Excluye campos de auditoría.
    
    Dim dictGlobal As Object
    
    On Error GoTo Errores
    
    Set dictGlobal = CreateObject("Scripting.Dictionary")
    
    ' 1. Datos Generales (tbSolicitudes)
    dictGlobal.Add "General", ObtenerDatosTabla("tbSolicitudes", "idSolicitud", idSolicitud)
    
    ' 2. Datos PC (tbDatosPC)
    dictGlobal.Add "PC", ObtenerDatosTabla("tbDatosPC", "idSolicitud", idSolicitud)
    
    ' 3. Datos CDCA (tbDatosCDCA)
    dictGlobal.Add "CDCA", ObtenerDatosTabla("tbDatosCDCA", "idSolicitud", idSolicitud)
    
    ' 4. Datos CDCASUB (tbDatosCDCASUB)
    dictGlobal.Add "CDCASUB", ObtenerDatosTabla("tbDatosCDCASUB", "idSolicitud", idSolicitud)
    
    ' Convertir a JSON
    GenerarSnapshotGlobal = JsonConverter.ConvertToJson(dictGlobal)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SnapshotServicio.GenerarSnapshotGlobal"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Function

Public Function CalcularHashSnapshot(ByVal idSolicitud As Long) As String
    Dim jsonSnapshot As String
    jsonSnapshot = GenerarSnapshotGlobal(idSolicitud)
    CalcularHashSnapshot = HashHelper.CalcularSHA256(jsonSnapshot)
End Function

Public Function ObtenerDatosTabla(ByVal nombreTabla As String, ByVal campoID As String, ByVal valorID As Long, Optional ByRef db As DAO.Database) As Object
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim fld As DAO.Field
    Dim dictFila As Object
    Dim sql As String

    Set dictFila = CreateObject("Scripting.Dictionary")

    On Error GoTo Errores
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    sql = "SELECT * FROM " & nombreTabla & " WHERE " & campoID & " = " & valorID
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenSnapshot)

    If Not rcd.EOF Then
        For Each fld In rcd.Fields
            If Not EsCampoExcluido(fld.name) Then
                ' Tratar Nulos como cadena vacía para consistencia
                dictFila.Add fld.name, "" & Nz(fld.value, "")
            End If
        Next fld
    End If

    rcd.Close
    Set ObtenerDatosTabla = dictFila
    Exit Function

Errores:
    ' Si la tabla no tiene datos para esa solicitud, devolvemos diccionario vacío (correcto)
    Set ObtenerDatosTabla = dictFila
End Function

Private Function EsCampoExcluido(ByVal nombreCampo As String) As Boolean
    ' Lista negra de campos de auditoría (Spec-105)
    Select Case LCase(nombreCampo)
        Case "fechacreacion", "fechamodificacion", _
             "usuariocreacion", "usuariomodificacion", _
             "idusuariocreacion", "idusuariomodificacion"
            EsCampoExcluido = True
        Case Else
            EsCampoExcluido = False
    End Select
End Function


