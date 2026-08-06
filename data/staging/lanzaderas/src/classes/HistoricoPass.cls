VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HistoricoPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public usuario As String
Public PassAntigua As String
Public FechaPass As String
Private m_objColCampos As Collection
Public Error As String

Public Function Save( _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim rcdDatos As DAO.Recordset
    
    On Error GoTo errores
    If Me.PassAntigua = "" Then
        p_Error = "No se ha indicado la pass"
        Err.Raise 1000
    End If
    If InStr(1, Me.usuario, "@") = 0 Then
        p_Error = "No se ha indicado el correo"
        Err.Raise 1000
    End If
    If Not IsDate(Me.FechaPass) Then
        Me.FechaPass = Now()
    End If
    
    m_SQL = "SELECT TbUsuariosHistoricoContrasenias.* " & _
            "FROM TbUsuariosHistoricoContrasenias " & _
            "WHERE (((TbUsuariosHistoricoContrasenias.Usuario)='" & Me.usuario & _
            "') AND ((TbUsuariosHistoricoContrasenias.PassAntigua)='" & Me.PassAntigua & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                .Fields("Usuario") = Me.usuario
                .Fields("PassAntigua") = Me.PassAntigua
                .Fields("FechaPass") = Me.FechaPass
            .Update
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Save = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HistoricoPass.Save ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function


Public Property Get ColCampos() As Collection
    
    Dim tdf As DAO.TableDef
    Dim fld As Object
    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set tdf = getdb().TableDefs("TbUsuariosHistoricoContrasenias")
    For Each fld In tdf.Fields
        If m_objColCampos Is Nothing Then
            Set m_objColCampos = New Collection
        
        End If
        m_objColCampos.Add fld.Name
    Next
    Set tdf = Nothing
    
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El metodo HistoricoPass.ColCampos ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "Usuario" Then
        getPropiedad = Me.usuario
        Exit Function
    End If
    If m_NombrePropiedad = "PassAntigua" Then
        getPropiedad = Me.PassAntigua
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPass" Then
        getPropiedad = Me.FechaPass
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en HistoricoPass.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo HistoricoPass.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "Usuario" Then
        Me.usuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PassAntigua" Then
        Me.PassAntigua = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPass" Then
        Me.FechaPass = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en HistoricoPass.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo HistoricoPass.SetPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function









