Attribute VB_Name = "RepositorioComun"

Option Compare Database
Option Explicit


Public Sub EjecutarAccion(ByVal sql As String, Optional params As Object, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim key As Variant
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    If dbConexion Is Nothing Then
        Dim errConn As New CondorError
        errConn.Create 513, "Conexion a base de datos no inicializada.", "RepositorioComun.EjecutarAccion"
        errConn.Raise
    End If
    
    Set qdf = dbConexion.CreateQueryDef("", sql)
    
    If Not params Is Nothing Then
        For Each key In params.Keys
            qdf.Parameters(key).value = params(key)
        Next key
    End If
    
    qdf.Execute dbFailOnError

LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RepositorioComun.EjecutarAccion"
    errObj.AddToCallStack "SQL: " & sql
    errObj.Raise
End Sub

Public Function EjecutarConsulta(ByVal sql As String, Optional params As Object, Optional db As DAO.Database) As DAO.Recordset
    Dim dbConexion As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim key As Variant
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    If dbConexion Is Nothing Then
        Dim errConn2 As New CondorError
        errConn2.Create 513, "Conexion a base de datos no inicializada.", "RepositorioComun.EjecutarConsulta"
        errConn2.Raise
    End If
    
    Set qdf = dbConexion.CreateQueryDef("", sql)
    
    If Not params Is Nothing Then
        For Each key In params.Keys
            qdf.Parameters(key).value = params(key)
        Next key
    End If
    
    Set EjecutarConsulta = qdf.OpenRecordset(dbOpenSnapshot)
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RepositorioComun.EjecutarConsulta"
    errObj.Raise
End Function

Public Function HidratarColeccionDesdeSQL(ByVal sql As String, ByVal nombreClase As String, _
                                        ByVal nombreCampoID As String, _
                                        Optional db As DAO.Database, _
                                        Optional params As Object) As Scripting.Dictionary
    Dim rcd As DAO.Recordset
    Dim objGenerico As Object
    Dim col As New Scripting.Dictionary
    On Error GoTo Errores
    
    col.CompareMode = TextCompare
    
    Set rcd = EjecutarConsulta(sql, params, db)
    
    If Not rcd.EOF Then
        rcd.MoveFirst
        Do While Not rcd.EOF
            Set objGenerico = Factoria.CreateEntity(nombreClase)
            If objGenerico Is Nothing Then
                Dim errFac As New CondorError
                errFac.Create 513, "La clase '" & nombreClase & "' no está registrada en la Factoria.", "RepositorioComun.HidratarColeccionDesdeSQL"
                errFac.Raise
            End If
            
            Call RellenarObjetoDesdeRecordset(objGenerico, rcd)
            col.Add CStr(objGenerico.getPropiedad(nombreCampoID)), objGenerico
            
            rcd.MoveNext
        Loop
    End If
    
    Set HidratarColeccionDesdeSQL = col
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RepositorioComun.HidratarColeccionDesdeSQL"
    errObj.AddToCallStack "Clase a hidratar: " & nombreClase
    errObj.Raise
End Function

Public Function HidratarEntidadDesdeSQL(ByVal sql As String, ByVal nombreClase As String, Optional db As DAO.Database, Optional params As Object) As Object
    Dim rcd As DAO.Recordset
    Dim objGenerico As Object
    On Error GoTo Errores
    
    Set rcd = EjecutarConsulta(sql, params, db)
    
    If Not rcd.EOF Then
        Set objGenerico = Factoria.CreateEntity(nombreClase)
        If objGenerico Is Nothing Then
             Err.Raise 513, "RepositorioComun.HidratarEntidad", "La clase '" & nombreClase & "' no está registrada en la Factoria."
        End If
        
        Call RellenarObjetoDesdeRecordset(objGenerico, rcd)
        Set HidratarEntidadDesdeSQL = objGenerico
    Else
        Set HidratarEntidadDesdeSQL = Nothing
    End If
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RepositorioComun.HidratarEntidadDesdeSQL"
    errObj.AddToCallStack "Clase a hidratar: " & nombreClase
    errObj.Raise
End Function


Private Sub RellenarObjetoDesdeRecordset(ByRef objDestino As Object, ByVal rcdFuente As DAO.Recordset)
    Dim campo As Variant
    Dim valor As Variant
    On Error GoTo Errores
    
    For Each campo In objDestino.ColCampos
        ' --- INICIO DE LA CORRECCIÓN ---
        ' Usamos Nz() con un solo argumento.
        ' Dejará que VBA decida el valor por defecto correcto ("" para strings, 0 para números/fechas)
        ' basándose en el tipo de la propiedad de destino.
        valor = Nz(rcdFuente.Fields(CStr(campo)).value)
        ' --- FIN DE LA CORRECCIÓN ---
        
        objDestino.SetPropiedad CStr(campo), valor
        
    Next campo
    
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RepositorioComun.RellenarObjetoDesdeRecordset"
    ' Enriquecemos el error con más contexto
    errObj.AddToCallStack "Clase: '" & TypeName(objDestino) & "', Campo: '" & CStr(campo) & "', Valor a Asignar: '" & valor & "'"
    errObj.Raise
End Sub


Public Sub RellenarRecordsetDesdeObjeto(ByRef rcdDestino As DAO.Recordset, ByRef objFuente As Object, Optional ByVal omitirPK As Boolean = False)
    Dim campo As Variant
    Dim nombreCampoPK As String
    
    On Error GoTo Errores
    
    ' El primer campo en ColCampos es, por convención, la clave primaria.
    If omitirPK Then
        nombreCampoPK = CStr(objFuente.ColCampos(1))
    End If
    
    For Each campo In objFuente.ColCampos
        If omitirPK And CStr(campo) = nombreCampoPK Then
            ' Si estamos en modo omisión y este es el campo PK, lo saltamos.
        Else
            rcdDestino.Fields(CStr(campo)).value = objFuente.getPropiedad(CStr(campo))
        End If
    Next campo
    
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RepositorioComun.RellenarRecordsetDesdeObjeto"
    errObj.AddToCallStack "Error al asignar el campo '" & CStr(campo) & "' a la clase '" & TypeName(objFuente) & "'"
    errObj.Raise
End Sub



