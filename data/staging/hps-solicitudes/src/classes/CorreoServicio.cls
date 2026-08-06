VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorreoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private m_Error As String

Public Function ObtenerCorreosGestion( _
                                    p_Estado As String, _
                                    p_IDSolicitud As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    On Error GoTo errores
    
    ' Aquí podríamos validar reglas de negocio si fuera necesario
    ' Por ejemplo, si el usuario tiene permiso para ver correos de esa solicitud
    
    ' Delegamos la consulta al repositorio
    Set ObtenerCorreosGestion = CorreoRepositorio.BuscarCorreos(p_Estado, p_IDSolicitud, p_Error)
    
    Exit Function
errores:
    p_Error = "Error en CorreoServicio.ObtenerCorreosGestion: " & Err.Description
End Function

