Attribute VB_Name = "NoConformidadRepositorio"


' ==========================================================================
' MÓDULO: NoConformidadRepositorio.bas (NUEVO REPOSITORIO)
' ==========================================================================
Option Compare Database
Option Explicit

Public Function getTodas(Optional ByRef db As DAO.Database = Nothing) As Scripting.Dictionary
    Dim sql As String
    Dim dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT IDNoConformidad, CodigoNoConformidad, Descripcion,idExpediente,CodConcesionAsociada  " & _
    "FROM TbNoConformidades " & _
    "ORDER BY IDNoConformidad DESC;"
    
    ' Usamos nuestra nueva función de conexión
    Set getTodas = RepositorioComun.HidratarColeccionDesdeSQL(sql, "NoConformidad", "IDNoConformidad", dbConexion)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadRepositorio.getTodas"
    errObj.Raise
End Function
Public Function getPorExpediente(ByVal idExp As Long, Optional ByRef db As DAO.Database = Nothing) As Scripting.Dictionary
    Dim sql As String
    Dim params As Object
    Dim dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' Filtramos por idExpediente. Asumimos que si idExpediente es 0 o Null en la tabla,
    ' no pertenece a este proyecto específico.
    sql = "SELECT IDNoConformidad, CodigoNoConformidad, Descripcion, idExpediente,CodConcesionAsociada " & _
          "FROM TbNoConformidades " & _
          "WHERE idExpediente = [p_IDExp] " & _
          "ORDER BY IDNoConformidad DESC;"
          
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDExp", idExp
    
    ' Usamos la conexión específica de NoConformidades
    Set getPorExpediente = RepositorioComun.HidratarColeccionDesdeSQL(sql, "NoConformidad", "IDNoConformidad", dbConexion, params)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadRepositorio.getPorExpediente"
    errObj.AddToCallStack "idExpediente: " & idExp
    errObj.Raise
End Function

Public Function getPorCodigoCondor(ByVal codigoCondor As String, Optional ByRef db As DAO.Database = Nothing) As NoConformidad
    Dim sql As String
    Dim rs As DAO.Recordset
    Dim nc As NoConformidad
    Dim params As Object
    Dim dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT IDNoConformidad, CodigoNoConformidad, CodConcesionAsociada, Descripcion, idExpediente " & _
          "FROM TbNoConformidades " & _
          "WHERE CodConcesionAsociada = [p_codigo] " & _
          "ORDER BY IDNoConformidad DESC;"
    
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_codigo", codigoCondor
    
    Set rs = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    If Not rs.EOF Then
        Set nc = New NoConformidad
        nc.IDNoConformidad = rs!IDNoConformidad
        nc.CodigoNoConformidad = Nz(rs!CodigoNoConformidad, "")
        nc.CodConcesionAsociada = Nz(rs!CodConcesionAsociada, "")
        nc.descripcion = Nz(rs!descripcion, "")
        nc.idExpediente = Nz(rs!idExpediente, 0)
        Set getPorCodigoCondor = nc
    End If
    rs.Close
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadRepositorio.getPorCodigoCondor"
    errObj.Raise
End Function

Public Function existeEnBaseDatosExternaa(ByVal idNC As Long, Optional ByRef db As DAO.Database = Nothing) As Boolean
    Dim rs As DAO.Recordset
    Dim dbConexion As DAO.Database
    Dim qdf As DAO.QueryDef
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    Set qdf = dbConexion.CreateQueryDef("", "SELECT COUNT(*) AS total FROM TbNoConformidades WHERE IDNoConformidad = ?")
    qdf.Parameters(0) = idNC
    Set rs = qdf.OpenRecordset()
    If Not rs.EOF Then
        existeEnBaseDatosExternaa = (rs!total > 0)
    End If
    rs.Close
    Set rs = Nothing
    qdf.Close
    Set qdf = Nothing
    
    Exit Function
Errores:
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    If Not qdf Is Nothing Then qdf.Close: Set qdf = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadRepositorio.existeEnBaseDatosExternaa"
    errObj.AddToCallStack "idNC: " & idNC
    errObj.Raise
End Function

Public Function actualizarCodigoConcesion(ByVal idNC As Long, ByVal codigoCondor As String, Optional ByRef dbExt As DAO.Database = Nothing) As Boolean
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores
    
    If dbExt Is Nothing Then Set dbExt = getdb()
    
    sql = "UPDATE TbNoConformidades SET CodConcesionAsociada = [p_codigo] " & _
          "WHERE IDNoConformidad = [p_id]"
    
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_codigo", codigoCondor
    params.Add "p_id", idNC
    
    Call RepositorioComun.EjecutarAccion(sql, params, dbExt)
    actualizarCodigoConcesion = True
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadRepositorio.actualizarCodigoConcesion"
    errObj.AddToCallStack "idNC: " & idNC & ", codigoCondor: " & codigoCondor
    errObj.Raise
End Function

Public Function limpiarCodigoConcesion(ByVal idNC As Long, Optional ByRef dbExt As DAO.Database = Nothing) As Boolean
    Dim qdf As DAO.QueryDef
    On Error GoTo Errores
    
    If dbExt Is Nothing Then Set dbExt = getdb()
    
    Set qdf = dbExt.CreateQueryDef("", "UPDATE TbNoConformidades SET CodConcesionAsociada = NULL WHERE IDNoConformidad = ?")
    qdf.Parameters(0) = idNC
    qdf.Execute
    limpiarCodigoConcesion = True
    qdf.Close
    Set qdf = Nothing
    
    Exit Function
Errores:
    If Not qdf Is Nothing Then qdf.Close: Set qdf = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadRepositorio.limpiarCodigoConcesion"
    errObj.AddToCallStack "idNC: " & idNC
    errObj.Raise
End Function

Public Function getPorID(ByVal idNC As Long, Optional ByRef db As DAO.Database = Nothing) As NoConformidad
    Dim sql As String
    Dim rs As DAO.Recordset
    Dim nc As NoConformidad
    Dim dbConexion As DAO.Database
    Dim params As Object
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT IDNoConformidad, CodigoNoConformidad, CodConcesionAsociada, Descripcion, idExpediente " & _
          "FROM TbNoConformidades " & _
          "WHERE IDNoConformidad = [p_id]"
    
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_id", idNC
    
    Set rs = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    If Not rs.EOF Then
        Set nc = New NoConformidad
        nc.IDNoConformidad = rs!IDNoConformidad
        nc.CodigoNoConformidad = Nz(rs!CodigoNoConformidad, "")
        nc.CodConcesionAsociada = Nz(rs!CodConcesionAsociada, "")
        nc.descripcion = Nz(rs!descripcion, "")
        nc.idExpediente = Nz(rs!idExpediente, 0)
        Set getPorID = nc
    End If
    rs.Close
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadRepositorio.getPorID"
    errObj.Raise
End Function

Public Sub LimpiarCodConcesionAsociada(ByVal codigoSolicitud As String, Optional ByRef dbExt As DAO.Database = Nothing)
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores
    
    If dbExt Is Nothing Then Set dbExt = getdb()
    
    sql = "UPDATE TbNoConformidades SET CodConcesionAsociada = NULL WHERE CodConcesionAsociada = [p_codigo]"
    
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_codigo", codigoSolicitud
    
    Call RepositorioComun.EjecutarAccion(sql, params, dbExt)
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadRepositorio.LimpiarCodConcesionAsociada"
    errObj.AddToCallStack "codigoSolicitud: " & codigoSolicitud
    errObj.Raise
End Sub






