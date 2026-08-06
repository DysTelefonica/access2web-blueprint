Attribute VB_Name = "MapeoRepositorio"

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: MapeoRepositorio.bas
' ==========================================================================

Public Function getMapeoParaPlantilla( _
    ByVal nombrePlantilla As String, _
    Optional ByRef p_Error As String = "" _
) As Object
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores
    
    sql = "SELECT * FROM tbMapeoCampos WHERE nombrePlantilla = [p_Plantilla];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_Plantilla", nombrePlantilla
    
    Dim db As DAO.Database
    Set db = getdb()
    If db Is Nothing Then
        p_Error = "getdb() returned Nothing"
        Set getMapeoParaPlantilla = Nothing
        Exit Function
    End If
    
    Set getMapeoParaPlantilla = RepositorioComun.HidratarColeccionDesdeSQL(sql, "MapeoCampos", "idMapeo", db, params)
    
    Exit Function
Errores:
    p_Error = Err.Number & ": " & Err.Description
    Set getMapeoParaPlantilla = Nothing
End Function

' ==========================================================================
' Fija drift de mapeo: actualiza nombreCampoTabla en tbMapeoCampos
' donde coincide plantilla + campoWord.
' Útil cuando cambia el nombre del campo en la tabla y hay que
' propagar el cambio a producción.
' Usa transacción explícita para no romper atomicidad.
' Error propagation via Optional ByRef p_Error As String.
' ==========================================================================
Public Function ActualizarCampoTabla( _
    ByVal nombrePlantilla As String, _
    ByVal nombreCampoWord As String, _
    ByVal nombreCampoTablaNuevo As String, _
    Optional ByRef p_FilasAfectadas As Long = 0, _
    Optional ByRef p_Error As String = "" _
) As Boolean
    Dim sql As String
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    On Error GoTo Errores

    Set db = getdb()
    If db Is Nothing Then
        p_Error = "getdb() returned Nothing"
        Exit Function
    End If

    Set ws = DBEngine.Workspaces(0)
    If ws Is Nothing Then
        p_Error = "DBEngine.Workspaces(0) returned Nothing"
        Exit Function
    End If

    ws.BeginTrans
    Dim qdf As DAO.QueryDef
    Set qdf = db.CreateQueryDef("")
    qdf.SQL = "PARAMETERS [p_CampoNuevo] Text, [p_Plantilla] Text, [p_CampoWord] Text; " & _
              "UPDATE tbMapeoCampos SET nombreCampoTabla = [p_CampoNuevo] " & _
              "WHERE nombrePlantilla = [p_Plantilla] AND nombreCampoWord = [p_CampoWord];"
    qdf.Parameters("p_CampoNuevo") = nombreCampoTablaNuevo
    qdf.Parameters("p_Plantilla") = nombrePlantilla
    qdf.Parameters("p_CampoWord") = nombreCampoWord
    qdf.Execute dbFailOnError
    p_FilasAfectadas = db.RecordsAffected
    Set qdf = Nothing
    ws.CommitTrans
    ActualizarCampoTabla = True

    Exit Function
Errores:
    p_Error = Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Rollback
    End If
    Set ws = Nothing
    Set db = Nothing
End Function