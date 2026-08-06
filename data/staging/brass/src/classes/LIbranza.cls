VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LIbranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDAusencia As String
Public Alias As String
Public FechaLibranza As String
Public Observaciones As String
Public MOTIVO As String
Public horas As String
Public ParaAlta As EnumSino
Private m_sIDAusenciaCalculada As String
Private m_ObjTecnico As Tecnico
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get IDAusenciaCalculada() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    m_sIDAusenciaCalculada = DameID("TbTecnicosAusencias", "IDAusencia")
    
    IDAusenciaCalculada = m_sIDAusenciaCalculada
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Libranza.IDAusenciaCalculada ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get Tecnico() As Tecnico

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjTecnico Is Nothing Then
        Set Tecnico = m_ObjTecnico
        Exit Property
    End If
    
    If Me.Alias = "" Then
        Exit Property
    End If
    Set m_ObjTecnico = Constructor.getTecnico(Me.Alias, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Tecnico = m_ObjTecnico
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Libranza.Tecnico ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property
Public Property Set Tecnico(m_Valor As Tecnico)

    Set m_ObjTecnico = m_Valor
End Property
Public Function MotivoNoOK( _
                                Optional ByRef p_ObjLibranzaAlInicio As LIbranza, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim m_IDAusencia As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Comprobar As EnumSino
    On Error GoTo errores
    
    If Me.ParaAlta = 0 Then
        MotivoNoOK = "No se sabe si es alta o edición"
        Err.Raise 1000
    End If
    If Me.ParaAlta = EnumSino.Sí Then
        If Me.IDAusencia = "" Then
            m_IDAusencia = Me.IDAusenciaCalculada
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not IsNumeric(m_IDAusencia) Then
                MotivoNoOK = "No se ha podido obtener el ID de la Ausencia"
                Exit Function
            End If
        End If
        
    Else
        If p_ObjLibranzaAlInicio Is Nothing Then
            If Me.IDAusencia = "" Then
                MotivoNoOK = "No se ha podido obtener el ID de la Ausencia"
                Exit Function
            End If
            Set p_ObjLibranzaAlInicio = Constructor.getLibranza(Me.IDAusencia, , p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If p_ObjLibranzaAlInicio Is Nothing Then
                MotivoNoOK = "No se ha podido obtener los datos de la ausencia"
                Exit Function
            End If
        End If
    End If
    If Me.Alias = "" Then
        MotivoNoOK = "Se ha de indicar el alias"
        Exit Function
    End If
    If Me.MOTIVO = "" Then
        MotivoNoOK = "Se ha de indicar el motivo"
        Exit Function
    End If
    If Not IsDate(Me.FechaLibranza) Then
        MotivoNoOK = "Se ha de indicar la fecha de la ausencia"
        Exit Function
    End If
    
    If Me.ParaAlta = EnumSino.Sí Then
        m_Comprobar = EnumSino.Sí
    Else
        If p_ObjLibranzaAlInicio.Alias <> Me.Alias Or p_ObjLibranzaAlInicio.FechaLibranza <> Me.FechaLibranza Then
            m_Comprobar = EnumSino.Sí
        End If
    End If
    If m_Comprobar = EnumSino.Sí Then
        m_SQL = "SELECT Count(TbTecnicosAusencias.IDAusencia) AS CuentaDeIDAusencia " & _
                "FROM TbTecnicosAusencias " & _
                "WHERE (((TbTecnicosAusencias.Alias)='" & Me.Alias & _
                "') AND ((TbTecnicosAusencias.FechaLibranza)=#" & Format(Me.FechaLibranza, "mm/dd/yyyy") & "#));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        If Not rcdDatos Is Nothing Then
            If IsNumeric(rcdDatos.Fields("CuentaDeIDAusencia")) Then
                If rcdDatos.Fields("CuentaDeIDAusencia") > 0 Then
                    rcdDatos.Close
                    Set rcdDatos = Nothing
                    MotivoNoOK = "La fecha ya ha sido registrada previamente"
                    Exit Function
                End If
            End If
        End If
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
       
    
    
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método Libranza.MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function



Public Function Registrar( _
                            Optional ByRef p_ObjLibranzaAlInicio As LIbranza, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim m_IDAusencia As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    
    m_MotivoNoOK = MotivoNoOK(p_ObjLibranzaAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.ParaAlta = EnumSino.Sí Then
        If Me.IDAusencia = "" Then
            m_IDAusencia = Me.IDAusenciaCalculada
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not IsNumeric(m_IDAusencia) Then
                p_Error = "No se ha podido obtener el ID de la Ausencia"
                Err.Raise 1000
            End If
            Me.IDAusencia = m_IDAusencia
        End If
        m_SQL = "TbTecnicosAusencias"
    Else
        m_SQL = "SELECT TbTecnicosAusencias.* " & _
                "FROM TbTecnicosAusencias " & _
                "WHERE (((TbTecnicosAusencias.Alias)='" & Me.Alias & _
                "') AND ((TbTecnicosAusencias.FechaLibranza)=#" & Format(Me.FechaLibranza, "mm/dd/yyyy") & "#));"
    End If
    Me.horas = 7.5
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Me.ParaAlta = EnumSino.Sí Then
            .AddNew
                .Fields("IDAusencia") = Me.IDAusencia
                .Fields("Alias") = Me.Alias
                .Fields("FechaLibranza") = Me.FechaLibranza
                If Me.Observaciones <> "" Then
                    .Fields("Observaciones") = Me.Observaciones
                End If
                .Fields("Motivo") = Me.MOTIVO
                .Fields("horas") = Me.horas
            .Update
        Else
            .Edit
                
                .Fields("Alias") = Me.Alias
                .Fields("FechaLibranza") = Me.FechaLibranza
                If Me.Observaciones <> "" Then
                    .Fields("Observaciones") = Me.Observaciones
                Else
                    .Fields("Observaciones") = Null
                End If
                .Fields("Motivo") = Me.MOTIVO
                .Fields("horas") = Me.horas
            .Update
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Libranza.Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                                
    On Error GoTo errores
    
    Me.Error = ""
    
    If Me.IDAusencia = "" Then
        p_Error = "No se ha indicado el ID"
        Err.Raise 1000
    End If
    m_SQL = "DELETE TbTecnicosAusencias.IDAusencia " & _
            "FROM TbTecnicosAusencias " & _
            "WHERE (((TbTecnicosAusencias.IDAusencia)=" & Me.IDAusencia & "));"
    CurrentDb().Execute m_SQL
   
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Libranza.Borrar ha devuelto el error: " & Err.Description
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
        .Add "IDAusencia"
        .Add "Alias"
        .Add "FechaLibranza"
        .Add "Observaciones"
        .Add "Motivo"
        .Add "horas"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Libranza.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDAusencia" Then
        getPropiedad = Me.IDAusencia
        Exit Function
    End If
    If m_NombrePropiedad = "Alias" Then
        getPropiedad = Me.Alias
        Exit Function
    End If
    If m_NombrePropiedad = "FechaLibranza" Then
        getPropiedad = Me.FechaLibranza
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    If m_NombrePropiedad = "Motivo" Then
        getPropiedad = Me.MOTIVO
        Exit Function
    End If
    If m_NombrePropiedad = "horas" Then
        getPropiedad = Me.horas
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Libranza.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Libranza.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDAusencia" Then
        Me.IDAusencia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Alias" Then
        Me.Alias = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaLibranza" Then
        Me.FechaLibranza = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Motivo" Then
        Me.MOTIVO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "horas" Then
        Me.horas = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Libranza.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Libranza.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


