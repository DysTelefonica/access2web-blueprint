VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fiesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public FechaFiesta As String
Public Observaciones As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Function MotivoNoOK( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_ValorExistente As EnumSino
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    If Not IsDate(Me.FechaFiesta) Then
        MotivoNoOK = "Se ha de indicar la fecha de la fiesta"
        Exit Function
    End If
    If Me.Observaciones = "" Then
        MotivoNoOK = "Se ha de indicar una fiesta"
        Exit Function
    End If
    m_SQL = "SELECT Count(TbTecnicosFiestas.FechaFiesta) AS CuentaDeFechaFiesta " & _
            "FROM TbTecnicosFiestas " & _
            "WHERE (((TbTecnicosFiestas.FechaFiesta)=#" & Format(Me.FechaFiesta, "mm/dd/yyyy") & "#));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If Not rcdDatos Is Nothing Then
        If IsNumeric(rcdDatos.Fields("CuentaDeFechaFiesta")) Then
            If rcdDatos.Fields("CuentaDeFechaFiesta") > 0 Then
                rcdDatos.Close
                Set rcdDatos = Nothing
                MotivoNoOK = "La fecha ya ha sido registrada previamente"
                Exit Function
            End If
        End If
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método Fiesta.MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function



Public Function Registrar( _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    
    m_MotivoNoOK = MotivoNoOK(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "TbTecnicosFiestas"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("FechaFiesta") = Me.FechaFiesta
            .Fields("Observaciones") = Me.Observaciones
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Fiesta.Registrar ha devuelto el error: " & Err.Description
    End If
End Function



Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                                
    On Error GoTo errores
    
    Me.Error = ""
    If Not IsDate(Me.FechaFiesta) Then
        p_Error = "No se ha indicado la fecha a borrar"
        Err.Raise 1000
    End If
    m_SQL = "DELETE * FROM TbTecnicosFiestas " & _
            "WHERE FechaFiesta=#" & Format(Me.FechaFiesta, "mm/dd/yyyy") & "#;"
    CurrentDb().Execute m_SQL
   
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Fiesta.Borrar ha devuelto el error: " & Err.Description
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
        .Add "FechaFiesta"
        .Add "Observaciones"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Fiesta.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "FechaFiesta" Then
        getPropiedad = Me.FechaFiesta
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Fiesta.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Fiesta.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "FechaFiesta" Then
        Me.FechaFiesta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Fiesta.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Fiesta.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

