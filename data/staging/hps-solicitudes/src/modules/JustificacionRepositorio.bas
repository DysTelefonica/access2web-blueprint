Attribute VB_Name = "JustificacionRepositorio"
' Módulo: JustificacionRepositorio.bas
Option Compare Database
Option Explicit

' Módulo: JustificacionRepositorio.bas

' CAMBIO: Renombramos a ObtenerTodas y quitamos el "WHERE activa = True"
Public Function ObtenerTodas() As Collection
    Dim sql As String
    ' Traemos TODO (Activas e Inactivas)
    sql = "SELECT idjustificacion, titulo, descripcion, activa " & _
          "FROM TbJustificaciones " & _
          "ORDER BY titulo"
    
    Set ObtenerTodas = RepositorioComun.HidratarColeccionDesdeSQL(sql, "Justificacion")
End Function

Public Function ObtenerPorID(id As Long) As Justificacion
    Dim sql As String
    sql = "SELECT * FROM TbJustificaciones WHERE idjustificacion = " & id
    
    Set ObtenerPorID = RepositorioComun.HidratarEntidadDesdeSQL(sql, "Justificacion")
End Function



Public Function Guardar(ByRef p_Justificacion As Justificacion, Optional ByRef p_Error As String) As Boolean
    Dim qdf As DAO.QueryDef
    Dim strSQL As String
    
    On Error GoTo errores
    
    ' 1. Cálculo del ID si es nuevo (No autonumérico)
    If p_Justificacion.IDJustificacion = 0 Then
        p_Justificacion.IDJustificacion = ObtenerUltimoID() + 1
        If p_Error <> "" Then Exit Function
    End If
    
    ' 2. Definición de la consulta con parámetros
    strSQL = "INSERT INTO TbJustificaciones (IdJustificacion, titulo, descripcion, activa) " & _
             "VALUES ([pId], [ptitulo], [pdescripcion], [pactiva]);"
    
    Set qdf = getdb().CreateQueryDef("", strSQL)
    With qdf
        .Parameters("[pId]") = p_Justificacion.IDJustificacion
        .Parameters("[ptitulo]") = p_Justificacion.Titulo
        .Parameters("[pdescripcion]") = p_Justificacion.Descripcion
        .Parameters("[pactiva]") = p_Justificacion.activa
        
        .Execute dbFailOnError
    End With
    
    Guardar = True
    Exit Function

errores:
    p_Error = "Error en JustificacionRepositorio.Guardar: " & Err.Description
    Guardar = False
End Function

Public Sub Eliminar(id As Long)
    ' Hacemos borrado LÓGICO para no romper integridad referencial
    Dim sql As String
    sql = "UPDATE TbJustificaciones SET activa = False WHERE idjustificacion = " & id
    RepositorioComun.EjecutarSQL sql
End Sub
' Función para comprobar si existen hijos dependientes
Public Function EstaSiendoUsada(id As Long) As Boolean
    Dim sql As String
    Dim rs As DAO.Recordset
    
    ' Asumo que tu tabla de solicitudes se llama 'TbSolicitudes' y la FK 'idjustificacion'
    ' Ajusta el nombre de la tabla si es diferente.
    sql = "SELECT Count(*) as Total FROM TbSolicitudes WHERE idjustificacion = " & id
    
    Set rs = getdb().OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then
        EstaSiendoUsada = (rs!total > 0)
    End If
    rs.Close
End Function

' Módulo: JustificacionRepositorio.bas

Public Function ObtenerUltimoID() As Long
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT Max(idjustificacion) FROM TbJustificaciones"
    
    ' IMPORTANTE: Usamos getDB() para acceder al Back-End directamente
    Set rs = getdb().OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        ' Nz por si la tabla estuviera vacía (devolvería 0)
        ObtenerUltimoID = Nz(rs.Fields(0), 0)
    Else
        ObtenerUltimoID = 0
    End If
    
    rs.Close
    Set rs = Nothing
End Function
' JustificacionRepositorio.bas


