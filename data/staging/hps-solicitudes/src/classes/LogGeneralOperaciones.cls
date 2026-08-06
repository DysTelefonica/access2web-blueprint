VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogGeneralOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Log As LogGeneral

Public Function Registrar( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    With Me.Log
        If .Accion = "" Then
            p_Error = "La acción es obligatoria"
            Err.Raise 1000
        End If
        If .CorreosAutomaticos <> "Sí" And .CorreosAutomaticos <> "No" Then
            p_Error = "Se ha de indicar si está o no configurado para correos automáticos"
            Err.Raise 1000
        End If
        .IDLog = .IDLogCalculado
        .FECHA = Ahora
        .Usuario = m_ObjUsuarioConectado.UsuarioRed
        m_SQL = "TbLogsGeneral"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        rcdDatos.AddNew
            rcdDatos.Fields("IDLog") = .IDLog
            rcdDatos.Fields("Fecha") = .FECHA
            rcdDatos.Fields("Accion") = .Accion
            rcdDatos.Fields("Usuario") = .Usuario
            If .Descripcion <> "" Then
                rcdDatos.Fields("Descripcion") = .Descripcion
            End If
            If .DNI <> "" Then
                rcdDatos.Fields("DNI") = .DNI
            End If
            If .IDSolicitud <> "" Then
                rcdDatos.Fields("IDSolicitud") = .IDSolicitud
            End If
            rcdDatos.Fields("CorreosAutomaticos") = .CorreosAutomaticos
        rcdDatos.Update
    
        rcdDatos.Close
        Set rcdDatos = Nothing
        
        
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método LogOperaciones.Registrar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function


