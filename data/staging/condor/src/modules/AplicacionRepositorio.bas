Attribute VB_Name = "AplicacionRepositorio"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: AplicacionRepositorio.bas (REFACTORIZADO COMPLETO CON CondorError)
' ==========================================================================

Public Function getAplicacionesPermisos(ByVal p_CorreoUsuario As String, Optional ByRef db As DAO.Database) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjUsuarioAplicacionPermisos As UsuarioAplicacionPermisos
    Dim colPermisos As New Scripting.Dictionary
    Dim params As Object
    Dim dbLocal As DAO.Database
    
    On Error GoTo Errores
    
    If p_CorreoUsuario = "" Then
        Set getAplicacionesPermisos = colPermisos
        Exit Function
    End If
    
    m_SQL = "SELECT * FROM TbUsuariosAplicacionesPermisos WHERE CorreoUsuario=[p_Correo];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_Correo", p_CorreoUsuario
    If db Is Nothing Then Set dbLocal = getdb() Else Set dbLocal = db
    Set rcdDatos = RepositorioComun.EjecutarConsulta(m_SQL, params, dbLocal)
    
    colPermisos.CompareMode = TextCompare
    
    If Not rcdDatos.EOF Then
        rcdDatos.MoveFirst
        Do While Not rcdDatos.EOF
            Set m_ObjUsuarioAplicacionPermisos = New UsuarioAplicacionPermisos
            For Each m_Campo In m_ObjUsuarioAplicacionPermisos.ColCampos
                
                m_ObjUsuarioAplicacionPermisos.SetPropiedad m_Campo, Nz(rcdDatos.Fields(m_Campo).value, "")
            Next
            
            If Not colPermisos.Exists(CStr(m_ObjUsuarioAplicacionPermisos.IDAplicacion)) Then
                colPermisos.Add m_ObjUsuarioAplicacionPermisos.IDAplicacion, m_ObjUsuarioAplicacionPermisos
            End If
            Set m_ObjUsuarioAplicacionPermisos = Nothing
            rcdDatos.MoveNext
        Loop
    End If

    Set getAplicacionesPermisos = colPermisos
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcdDatos Is Nothing Then rcdDatos.Close
    Exit Function
    
Errores:
    If Not rcdDatos Is Nothing Then rcdDatos.Close
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AplicacionRepositorio.getAplicacionesPermisos"
    errObj.Raise
End Function
