VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TipoTecnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Tipo As String
Public ParaAlta As EnumSino

Private m_sPrecioHoraLabActual As String
Private m_sPrecioHoraExtraActual As String

Private m_objColTecnicos As Scripting.Dictionary
Private m_objColTecnicosActivos As Scripting.Dictionary
Private m_objColTecnicosNOActivos As Scripting.Dictionary
Private m_objColCampos As Collection

Public Error As String
Private m_Error As String

Public Property Get ColTecnicos() As Scripting.Dictionary
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColTecnicos Is Nothing Then
        Set ColTecnicos = m_objColTecnicos
        Exit Property
    End If
    Set m_objColTecnicos = Constructor.getTecnicosPorTipo(Me.Tipo, 0, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColTecnicos = m_objColTecnicos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnico.ColTecnicos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColTecnicos(ByVal objNewValue As Variant)

    Set m_objColTecnicos = objNewValue

End Property

Public Property Get ColTecnicosActivos() As Scripting.Dictionary
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColTecnicosActivos Is Nothing Then
        Set ColTecnicosActivos = m_objColTecnicosActivos
        Exit Property
    End If
    Set m_objColTecnicosActivos = Constructor.getTecnicosPorTipo(Me.Tipo, EnumSino.Sí, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColTecnicosActivos = m_objColTecnicosActivos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnico.ColTecnicosActivos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColTecnicosActivos(ByVal objNewValue As Variant)

    Set m_objColTecnicosActivos = objNewValue

End Property

Public Property Get ColTecnicosNOActivos() As Scripting.Dictionary
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColTecnicosNOActivos Is Nothing Then
        Set ColTecnicosNOActivos = m_objColTecnicosNOActivos
        Exit Property
    End If
    Set m_objColTecnicosNOActivos = Constructor.getTecnicosPorTipo(Me.Tipo, EnumSino.No, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColTecnicosNOActivos = m_objColTecnicosNOActivos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnico.ColTecnicosNOActivos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColTecnicosNOActivos(ByVal objNewValue As Variant)

    Set m_objColTecnicosNOActivos = objNewValue

End Property

Public Property Get PrecioHoraLabActual() As String

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If IsNumeric(m_sPrecioHoraLabActual) Then
        PrecioHoraLabActual = m_sPrecioHoraLabActual
        Exit Property
    End If
    If Me.Tipo = "" Then
        Exit Property
    End If
    m_sPrecioHoraLabActual = getPrecioHora(Me.Tipo, EnumSino.Sí, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PrecioHoraLabActual = m_sPrecioHoraLabActual
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnico.PrecioHoraLabActual ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Get PrecioHoraExtraActual() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If IsNumeric(m_sPrecioHoraExtraActual) Then
        PrecioHoraExtraActual = m_sPrecioHoraExtraActual
        Exit Property
    End If
    If Me.Tipo = "" Then
        Exit Property
    End If
    m_sPrecioHoraExtraActual = getPrecioHora(Me.Tipo, EnumSino.No, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PrecioHoraExtraActual = m_sPrecioHoraExtraActual
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnico.PrecioHoraExtraActual ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Private Function getPrecioHora( _
                                p_Tipo As String, _
                                Optional p_EsLaborable As EnumSino = EnumSino.Sí, _
                                Optional p_Error As String _
                                ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_NombreCampo As String
    
    On Error GoTo errores
    If p_Tipo = "" Then
        Exit Function
    End If
    If p_EsLaborable = EnumSino.Sí Then
        m_NombreCampo = "PrecioHoraLab"
    Else
        m_NombreCampo = "PrecioHoraExtra"
    End If
    m_SQL = "SELECT " & m_NombreCampo & " " & _
            "FROM TbTipoTecnicoPrecios " & _
            "WHERE TIPOTECNICO='" & p_Tipo & "' " & _
            "ORDER BY IDTipoTecnicoFecha DESC;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(.Fields(m_NombreCampo)) Then
                getPrecioHora = .Fields(m_NombreCampo)
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método Tecnico.getPrecioHora ha devuelto el error: " & Err.Description
     End If
End Function

Public Function MotivoNoOK( _
                                Optional ByRef p_ObjTipoTecnicoAlInicio As TIPOTECNICO, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_ValorExistente As EnumSino
    Dim m_TipoEnAlgunaFactura As EnumSino
    On Error GoTo errores
        
    With Me
        If .ParaAlta = 0 Then
            MotivoNoOK = "No se sabe si es alta o edición"
            Exit Function
            
        End If
        If .Tipo = "" Then
            MotivoNoOK = "Obligatorio el Tipo"
            Exit Function
        End If
        
        If Not p_ObjTipoTecnicoAlInicio Is Nothing Then
            If p_ObjTipoTecnicoAlInicio.Tipo <> Me.Tipo Then
                m_ValorExistente = ValorExistente("TbTipoTecnico", "TIPO", Me.Tipo, EnumSino.No, CurrentDb, p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If m_ValorExistente = EnumSino.Sí Then
                    MotivoNoOK = "El TIPO ya ha sido usado"
                    Exit Function
                End If
                m_TipoEnAlgunaFactura = TipoEnAlgunaFactura(Me.Tipo, p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If m_TipoEnAlgunaFactura = EnumSino.Sí Then
                    MotivoNoOK = "El TIPO ya ha sido usado en alguna factura"
                    Exit Function
                End If
            End If
            
        Else
           m_ValorExistente = ValorExistente("TbTipoTecnico", "TIPO", Me.Tipo, EnumSino.No, CurrentDb, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ValorExistente = EnumSino.Sí Then
                MotivoNoOK = "El TIPO ya ha sido usado"
                Exit Function
            End If
        End If
        
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método Tecnico.MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function

Private Function TipoEnAlgunaFactura( _
                                        p_Tipo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As EnumSino
            
    Dim rcdDatos As DAO.Recordset
    
    On Error GoTo errores
        
    m_SQL = "SELECT Count(TbFacturaEventosInvolucrados.IDFactura) AS CuentaDeIDFactura " & _
            "FROM TbFacturaEventosInvolucrados " & _
            "WHERE (((TbFacturaEventosInvolucrados.TIPOTECNICO)='" & p_Tipo & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    If Not rcdDatos Is Nothing Then
        If IsNumeric(Nz(rcdDatos.Fields("CuentaDeIDFactura"), "")) Then
            If CLng(rcdDatos.Fields("CuentaDeIDFactura")) > 0 Then
                TipoEnAlgunaFactura = EnumSino.Sí
                rcdDatos.Close
                Set rcdDatos = Nothing
            End If
        End If
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    m_SQL = "SELECT Count(TbFacturaActividadesInvolucradas.IDFactura) AS CuentaDeIDFactura " & _
            "FROM TbFacturaActividadesInvolucradas " & _
            "WHERE (((TbFacturaActividadesInvolucradas.TIPOTECNICO)='" & p_Tipo & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    If Not rcdDatos Is Nothing Then
        If IsNumeric(Nz(rcdDatos.Fields("CuentaDeIDFactura"), "")) Then
            If CLng(rcdDatos.Fields("CuentaDeIDFactura")) > 0 Then
                TipoEnAlgunaFactura = EnumSino.Sí
                rcdDatos.Close
                Set rcdDatos = Nothing
            End If
        End If
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    TipoEnAlgunaFactura = EnumSino.No
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método TipoTecnico.TipoEnAlgunaFactura ha devuelto el error: " & Err.Description
     End If
                                
End Function

Public Function Registrar( _
                            Optional ByRef p_ObjTipoTecnicoAlInicio As TIPOTECNICO, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If Me.ParaAlta = 0 Then
        p_Error = "No se sabe si es alta o edición"
        Err.Raise 1000
    End If
    If Me.ParaAlta = EnumSino.No Then
        If p_ObjTipoTecnicoAlInicio Is Nothing Then
            Set p_ObjTipoTecnicoAlInicio = Constructor.getTipoTecnico(Me.Tipo, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    m_MotivoNoOK = MotivoNoOK(p_ObjTipoTecnicoAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.ParaAlta = EnumSino.Sí Then
        m_SQL = "TbTipoTecnico"
        
    Else
        m_SQL = "SELECT * FROM TbTipoTecnico " & _
            "WHERE TIPO='" & p_ObjTipoTecnicoAlInicio.Tipo & "';"
        
    End If
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Me.ParaAlta = EnumSino.No Then
            If .EOF Then
                p_Error = "No se ha encontrado el registro"
                Err.Raise 1000
            End If
        End If
        If Me.ParaAlta = EnumSino.Sí Then
            .AddNew
                .Fields("TIPO") = Me.Tipo
                
            .Update
        Else
            .Edit
                .Fields("TIPO") = Me.Tipo
            .Update
            m_SQL = "UPDATE TbTecnicos SET TbTecnicos.TIPO = '" & Me.Tipo & "' " & _
                    "WHERE (((TbTecnicos.TIPO)='" & p_ObjTipoTecnicoAlInicio.Tipo & "'));"
            CurrentDb().Execute m_SQL
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TipoTecnico.Registrar ha devuelto el error: " & Err.Description
    End If
End Function


Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    
    
    
    Dim m_ValorExistente As EnumSino
    Dim m_TipoEnAlgunaFactura As EnumSino
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_TipoEnAlgunaFactura = TipoEnAlgunaFactura(Me.Tipo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoEnAlgunaFactura = EnumSino.Sí Then
        p_Error = "El TIPO ya ha sido usado en alguna factura"
        Err.Raise 1000
    End If
    m_ValorExistente = ValorExistente("TbTecnicos", "TIPO", Me.Tipo, EnumSino.No, CurrentDb, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ValorExistente = EnumSino.Sí Then
        p_Error = "Al menos hay un usuario que está usando ese tipo"
        Exit Function
    End If

        
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TipoTecnico.ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                                
    On Error GoTo errores
    
    Me.Error = ""
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "DELETE * FROM TbTipoTecnico " & _
            "WHERE TIPO='" & Me.Tipo & "';"
    CurrentDb().Execute m_SQL
   
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TipoTecnico.Borrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Property Get ColCampos() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set m_objColCampos = New Collection
    With m_objColCampos
        .Add "TIPO"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoTecnico.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "TIPO" Then
        getPropiedad = Me.Tipo
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en TipoTecnico.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TipoTecnico.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "TIPO" Then
        Me.Tipo = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en TipoTecnico.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TipoTecnico.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


