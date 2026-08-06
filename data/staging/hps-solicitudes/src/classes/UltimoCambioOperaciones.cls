VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UltimoCambioOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public Function Registrar( _
                            p_Solicitud As solicitud, _
                            Optional ByRef p_Error As String _
                            ) As UltimoCambio

    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    Dim m_UltimoCambio As UltimoCambio
    
    On Error GoTo errores
    
    Set m_UltimoCambio = New UltimoCambio
    With m_UltimoCambio
        
        .IDSolicitud = p_Solicitud.IDSolicitud
        .IDUsuarioCambio = m_ObjUsuarioConectado.id
        .FechaCambio = Now()
        m_SQL = "SELECT * " & _
                "FROM TbUltimoCambio " & _
                "WHERE IDUsuarioCambio=" & .IDUsuarioCambio & ";"
        
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        If rcdDatos.EOF Then
            rcdDatos.AddNew
                .id = .IDCalculado
                rcdDatos.Fields("ID") = CLng(.id)
                rcdDatos.Fields("IDSolicitud") = .IDSolicitud
                rcdDatos.Fields("FechaCambio") = .FechaCambio
                rcdDatos.Fields("IDUsuarioCambio") = .IDUsuarioCambio
            rcdDatos.Update
        Else
            rcdDatos.Edit
                rcdDatos.Fields("IDSolicitud") = .IDSolicitud
                rcdDatos.Fields("FechaCambio") = .FechaCambio
                rcdDatos.Fields("IDUsuarioCambio") = .IDUsuarioCambio
            rcdDatos.Update
        End If
        
        rcdDatos.Close
        Set rcdDatos = Nothing
    End With
    
    With p_Solicitud
        m_SQL = "SELECT * " & _
                "FROM TbSolicitudes " & _
                "WHERE IDSolicitud=" & .IDSolicitud & ";"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        If Not rcdDatos.EOF Then
            rcdDatos.Edit
                If Not IsDate(.FechaCreacion) Then
                    .FechaCreacion = Now()
                    .UsuarioCreacion = m_ObjUsuarioConectado.UsuarioRed
                    rcdDatos.Fields("FechaCreacion") = .FechaCreacion
                    rcdDatos.Fields("UsuarioCreacion") = .UsuarioCreacion
                Else
                    .FechaUltimoCambio = Now()
                    .UsuarioUltimoCambio = m_ObjUsuarioConectado.UsuarioRed
                    rcdDatos.Fields("FechaUltimoCambio") = .FechaUltimoCambio
                    rcdDatos.Fields("UsuarioUltimoCambio") = .UsuarioUltimoCambio
                End If
            rcdDatos.Update
        End If
        
        rcdDatos.Close
        Set rcdDatos = Nothing
        
    End With
    Set Registrar = m_UltimoCambio
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UltimoCambioOperaciones.Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function



