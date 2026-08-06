VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResponsableOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Responsable As Responsable

Public Function Registrar( _
                            Optional p_ResponsableAlInicio As Responsable, _
                            Optional ByRef p_Error As String _
                            ) As String

    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    Dim m_Responsable As Responsable
    Dim m_Comprobar As Boolean
    
    On Error GoTo errores
    With Me.Responsable
        If .Nombre = "" Then
            p_Error = "Falta el campo Nombre"
            Err.Raise 1000
        End If
        If Not EsCorreoValido(.Correo) Then
            p_Error = "Correo no tiene un formato correcto de email"
            Err.Raise 1000
        End If
        If Not p_ResponsableAlInicio Is Nothing Then
            If .Correo <> p_ResponsableAlInicio.Correo Then
                m_Comprobar = True
            End If
        Else
            m_Comprobar = True
        End If
        If m_Comprobar = True Then
            Set m_Responsable = constructor.getResponsable(p_Correo:=.Correo, p_Error:=p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not m_Responsable Is Nothing Then
                p_Error = "Ese correo de Responsable ya está registrado"
                Err.Raise 1000
            End If
        End If
        
        
        If p_ResponsableAlInicio Is Nothing Then
            .IDResponsable = .IDResponsableCalculado
            m_SQL = "TbResponsables"
        Else
            m_SQL = "SELECT * " & _
                "FROM TbResponsables " & _
                "WHERE IDResponsable=" & .IDResponsable & ";"
        End If
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        If p_ResponsableAlInicio Is Nothing Then
            rcdDatos.AddNew
                rcdDatos.Fields("IDResponsable") = .IDResponsable
                rcdDatos.Fields("Nombre") = .Nombre
                rcdDatos.Fields("Correo") = .Correo
                If .Descripcion <> "" Then
                    rcdDatos.Fields("Descripcion") = .Descripcion
                End If
                
                
            rcdDatos.Update
        Else
            rcdDatos.Edit
                rcdDatos.Fields("Nombre") = .Nombre
                rcdDatos.Fields("Correo") = .Correo
                If .Descripcion <> "" Then
                    rcdDatos.Fields("Descripcion") = .Descripcion
                Else
                    rcdDatos.Fields("Descripcion") = Null
                End If
            rcdDatos.Update
        End If
        rcdDatos.Close
        Set rcdDatos = Nothing
    End With
    Set m_ObjEntorno.Responsables = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ResponsableOperaciones.Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
   
    Dim m_ColSolicitudes As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Solicitud As solicitud
    
    Dim m_IDSolicitud As String
    On Error GoTo errores
    
    Set m_ColSolicitudes = getSolicitudesConResponsable(p_Correo:=Me.Responsable.Correo, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColSolicitudes Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ColSolicitudes
        Set m_Solicitud = m_ColSolicitudes(m_ID)
        If m_IDSolicitud = "" Then
            m_IDSolicitud = m_Solicitud.IDSolicitud
        Else
            m_IDSolicitud = m_IDSolicitud & vbNewLine & m_Solicitud.IDSolicitud
        End If
        Set m_Solicitud = Nothing
    Next
    p_Error = "El correo de ese responsable está siendo usado en las solicitudes: " & vbNewLine & m_IDSolicitud
    Err.Raise 1000
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ResponsableOperaciones.ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Eliminar( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    With Me.Responsable
       
        
        m_SQL = "DELETE * " & _
                "FROM TbResponsables " & _
                "WHERE IDResponsable=" & .IDResponsable & ";"
        getdb().Execute m_SQL
        
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ResponsableOperaciones.Eliminar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
