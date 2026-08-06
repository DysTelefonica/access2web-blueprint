VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpedienteServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: ExpedienteServicio.cls (REFACTORIZADO CON CondorError)
' ==========================================================================

Public Function getExpedientes(Optional p_PalabraClave As String, _
                               Optional p_IDRespCalidad As String, _
                               Optional p_IDRespTecnico As String) As Scripting.Dictionary
    Dim sql As String, params As Object, wheres As Object
    Dim fromClause As String
    Dim idCalidad As Long, idTecnico As Long

    On Error GoTo Errores

    idCalidad = 0
    idTecnico = 0
    If IsNumeric(p_IDRespCalidad) Then idCalidad = CLng(p_IDRespCalidad)
    If IsNumeric(p_IDRespTecnico) Then idTecnico = CLng(p_IDRespTecnico)

    Set wheres = CreateObject("Scripting.Dictionary")
    Set params = CreateObject("Scripting.Dictionary")
    
    fromClause = "FROM (((TbExpedientes AS T_Exp " & _
                 "LEFT JOIN TbUsuariosAplicaciones AS T_RC ON T_Exp.IDResponsableCalidad = T_RC.Id) " & _
                 "LEFT JOIN (SELECT T_Resp_RT.IdExpediente, FIRST(T_Usr_RT.Nombre) AS NombreRT, FIRST(T_Usr_RT.CorreoUsuario) AS EmailRT " & _
                            "FROM (TbExpedientesResponsables AS T_Resp_RT INNER JOIN TbUsuariosAplicaciones AS T_Usr_RT ON T_Resp_RT.IdUsuario = T_Usr_RT.Id) " & _
                            "WHERE T_Resp_RT.EsJefeProyecto = 'Sí' GROUP BY T_Resp_RT.IdExpediente) AS SubRT ON T_Exp.IDExpediente = SubRT.IdExpediente) " & _
                 "LEFT JOIN (SELECT T_ExpRACS.IDEXPEDIENTE, FIRST(T_RACS.RAC) AS NombreRAC " & _
                            "FROM (TbExpedientesRACS AS T_ExpRACS INNER JOIN TbRACS AS T_RACS ON T_ExpRACS.IDRAC = T_RACS.IDRAC) " & _
                            "GROUP BY T_ExpRACS.IDEXPEDIENTE) AS SubRAC ON T_Exp.IDExpediente = SubRAC.IDEXPEDIENTE) "

    sql = "SELECT " & _
            "T_Exp.IDExpediente, T_Exp.Nemotecnico, T_Exp.Titulo,T_Exp.objetoContrato, T_Exp.CodExp,T_Exp.CodExpLargo,T_Exp.CodigoActividad, " & _
            "T_Exp.ContratistaPrincipal, T_Exp.FechaInicioContrato, T_Exp.FechaFinContrato, T_Exp.FechaFinGarantia, " & _
            "T_RC.Nombre AS ResponsableCalidad, T_RC.CorreoUsuario AS EmailResponsableCalidad, " & _
            "SubRT.NombreRT AS ResponsableTecnico, " & _
            "SubRT.EmailRT AS EmailResponsableTecnico, " & _
            "SubRAC.NombreRAC AS RacAsignado " & _
          fromClause

    wheres.Add "(T_Exp.Pecal)='Sí'", ""

    If Nz(p_PalabraClave, "") <> "" Then
        wheres.Add "(T_Exp.Nemotecnico LIKE [p_PalabraClave] OR T_Exp.CodExp LIKE [p_PalabraClave])", ""
        params.Add "p_PalabraClave", "*" & p_PalabraClave & "*"
    End If
    
    If idCalidad > 0 Then
        wheres.Add "T_RC.Id = [p_IDRespCalidad]", ""
        params.Add "p_IDRespCalidad", idCalidad
    End If
    
    If idTecnico > 0 Then
        wheres.Add "EXISTS (" & _
                   "SELECT 1 FROM TbExpedientesResponsables AS T_Resp_Filtro " & _
                   "WHERE T_Resp_Filtro.IdExpediente = T_Exp.IDExpediente " & _
                   "AND T_Resp_Filtro.IdUsuario = [p_IDRespTecnico] " & _
                   "AND T_Resp_Filtro.EsJefeProyecto = 'Sí')", ""
        params.Add "p_IDRespTecnico", idTecnico
    End If

    If wheres.count > 0 Then
        sql = sql & " WHERE " & Join(wheres.Keys, " AND ")
    End If
    sql = sql & " ORDER BY T_Exp.IDExpediente DESC;"
    
    Set getExpedientes = RepositorioComun.HidratarColeccionDesdeSQL(sql, "Expediente", "IDExpediente", getdb(), params)

    Exit Function

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ExpedienteServicio.getExpedientes"
    errObj.Raise
End Function

Public Function getExpedientePorID(ByVal idExpediente As String) As Expediente
    On Error GoTo Errores
    
    If Trim(Nz(idExpediente, "")) = "" Or Not IsNumeric(idExpediente) Then
        Err.Raise 513, , "Se requiere un ID de Expediente válido."
    End If
    
    Set getExpedientePorID = ExpedienteRepositorio.getExpedientePorID(idExpediente)

    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ExpedienteServicio.getExpedientePorID"
    errObj.Raise
End Function

Public Function getClaveParaCodigoSolicitud(ByRef exp As Expediente) As String
    On Error GoTo Errores
    
    ' Regla a) Prioridad absoluta: Si hay Código de Actividad, se usa ese.
    If Trim(Nz(exp.CodigoActividad, "")) <> "" Then
        getClaveParaCodigoSolicitud = exp.CodigoActividad
        Exit Function
    End If
    
    ' Regla b) Analizar CodExp "sucio" (ej: "0395/20 (1559/23)")
    Dim codExpVal As String
    codExpVal = Trim(Nz(exp.CodExp, ""))
    
    ' Estrategia: Buscar la primera barra "/"
    ' Todo lo que hay antes (limpio) es el Expediente (XXXX)
    ' Los 2 caracteres inmediatamente posteriores son el Año (AA)
    If InStr(1, codExpVal, "/") > 0 Then
        Dim partes() As String
        partes = Split(codExpVal, "/")
        
        ' Verificamos que al menos haya una parte antes y una después de la barra
        If UBound(partes) >= 1 Then
            Dim strExpediente As String
            Dim strAnio As String
            
            ' Parte 1: Expediente (todo lo anterior a la primera barra)
            ' Ej: "0395" de "0395/20..."
            ' Usamos Split(0) y limpiamos espacios por si viene " 0395 "
            strExpediente = Trim(partes(0))
            
            ' Parte 2: Año (solo los 2 primeros caracteres después de la barra)
            ' Ej: De "20 (1559/23)", cogemos solo "20"
            strAnio = Left(Trim(partes(1)), 2)
            
            ' Validación mínima: El año debe ser numérico y el expediente tener contenido
            If Len(strExpediente) > 0 And IsNumeric(strAnio) Then
                ' Construimos según regla: AAXXXX
                ' 20 + 0395 = 200395
                getClaveParaCodigoSolicitud = strAnio & strExpediente
                Exit Function
            End If
        End If
    End If
    
    ' Fallback: Si no hay barra ni actividad, devolvemos ERROR o el valor limpio
    ' Esto permite al usuario verlo y quizás editarlo manualmente en lugar de crashear.
    Err.Raise 513, "Generación de ClaveExp", "No se pudo extraer el patrón XXXX/AA del expediente: " & exp.Nemotecnico

LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ExpedienteServicio.getClaveParaCodigoSolicitud"
    errObj.Raise
End Function
' --- FIN DE LA REFACTORIZACIÓN ---
Public Function getNombresSuministradores(ByVal idExpediente As Long, Optional ByRef db As DAO.Database) As Scripting.Dictionary
    Dim dictResult As New Scripting.Dictionary
    Dim colRelaciones As Scripting.Dictionary
    Dim key As Variant
    Dim relacion As ExpedienteSuministrador
    Dim suministrador As suministrador
    Dim sumServ As New SuministradorServicio
    
    ' Variables para construir las cadenas concatenadas
    Dim strPrincipales As String
    Dim strSubs As String
    
    On Error GoTo Errores
    
    ' Inicializamos las cadenas
    strPrincipales = ""
    strSubs = ""
    
    Set colRelaciones = ExpedienteRepositorio.getSuministradoresPorExpediente(idExpediente, db)
    
    If Not colRelaciones Is Nothing Then
        For Each key In colRelaciones.Keys
            Set relacion = colRelaciones(key)
            Set suministrador = sumServ.getSuministradorPorID(CStr(relacion.idSuministrador), db)
            
            If Not suministrador Is Nothing Then
                ' Lógica para Contratistas Principales
                If relacion.ContratistaPrincipal = "Sí" Then
                    If strPrincipales = "" Then
                        strPrincipales = suministrador.NombreParaMostrar
                    Else
                        strPrincipales = strPrincipales & " | " & suministrador.NombreParaMostrar
                    End If
                End If
                
                ' Lógica para Subcontratistas
                If relacion.SubContratista = "Sí" Then
                    If strSubs = "" Then
                        strSubs = suministrador.NombreParaMostrar
                    Else
                        strSubs = strSubs & " | " & suministrador.NombreParaMostrar
                    End If
                End If
            End If
        Next key
    End If
    
    ' Asignamos los resultados finales al diccionario
    dictResult.Add "principal", IIf(strPrincipales = "", "N/A", strPrincipales)
    dictResult.Add "subcontratista", IIf(strSubs = "", "N/A", strSubs)
    
    Set getNombresSuministradores = dictResult

LimpiarYSalir:
    On Error Resume Next
    Set dictResult = Nothing
    Set colRelaciones = Nothing
    Set relacion = Nothing
    Set suministrador = Nothing
    Set sumServ = Nothing
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ExpedienteServicio.getNombresSuministradores"
    errObj.Raise
End Function

Public Function getSuministradoresPorExpediente(ByVal idExpediente As Long, Optional ByRef db As DAO.Database) As Scripting.Dictionary
    ' Responsabilidad: Obtener la colección de relaciones Expediente-Suministrador.
    ' Delega la llamada de acceso a datos directamente al repositorio.
    On Error GoTo Errores
    
    ' Validación básica de la entrada
    If idExpediente <= 0 Then
        ' Devolvemos una colección vacía si el ID no es válido
        Set getSuministradoresPorExpediente = New Scripting.Dictionary
        Exit Function
    End If
    
    Set getSuministradoresPorExpediente = ExpedienteRepositorio.getSuministradoresPorExpediente(idExpediente, db)
    
LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ExpedienteServicio.getSuministradoresPorExpediente"
    errObj.AddToCallStack "ID Expediente: " & idExpediente
    errObj.Raise
End Function
Public Function getExpedienteViewModelPorID(ByVal idExpediente As String) As ExpedienteViewModel
    Dim vm As ExpedienteViewModel
    Dim exp As Expediente
    Dim dictNombres As Scripting.Dictionary
    
    On Error GoTo Errores
    
    ' 1. Obtener la entidad principal del expediente
    Set exp = Me.getExpedientePorID(idExpediente)
    If exp Is Nothing Then
        ' Si no se encuentra el expediente, devolvemos Nothing
        Set getExpedienteViewModelPorID = Nothing
        Exit Function
    End If
    
    ' 2. Construir el ViewModel
    Set vm = New ExpedienteViewModel
    Set vm.Expediente = exp
    
    ' 3. Obtener los datos relacionados y pre-calculados
    Set dictNombres = Me.getNombresSuministradores(exp.idExpediente)
    vm.NombresContratistas = dictNombres("principal")
    vm.NombresSubcontratistas = dictNombres("subcontratista")
    
    ' 4. Devolver el ViewModel completo
    Set getExpedienteViewModelPorID = vm

LimpiarYSalir:
    Set vm = Nothing
    Set exp = Nothing
    Set dictNombres = Nothing
    Exit Function
    
Errores:
    ' Propagamos el error usando nuestro sistema canónico
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ExpedienteServicio.getExpedienteViewModelPorID"
    errObj.AddToCallStack "idExpediente: " & idExpediente
    errObj.Raise
End Function

Public Function GetRacAsignadoPorExpediente(ByVal idExpediente As Long, Optional ByRef db As DAO.Database) As String
    On Error GoTo Errores
    
    GetRacAsignadoPorExpediente = ExpedienteRepositorio.GetRacAsignadoPorExpediente(idExpediente, db)

    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ExpedienteServicio.GetRacAsignadoPorExpediente"
    errObj.AddToCallStack "ID Expediente: " & idExpediente
    errObj.Raise
End Function



