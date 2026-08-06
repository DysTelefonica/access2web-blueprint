VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PCAccionReversa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public IDAccionContingencia As String
Public IDEdicion As String
Public FechaInicio As String
Public FechaFinPrevista As String
Public FechaFinReal As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Function Registrar( _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    If Me.IDAccionContingencia = "" Then
        p_Error = "Falta IDAccionContingencia"
        Err.Raise 1000
    End If
    If Me.IDEdicion = "" Then
        p_Error = "Falta IDEdicion"
        Err.Raise 1000
    End If
    m_SQL = "DELETE * " & _
            "FROM TbRiesgosPlanContingenciaDetalleReversa " & _
            "WHERE IDAccionContingencia=" & Me.IDAccionContingencia & ";"
    
    getdb().Execute m_SQL
    m_SQL = "TbRiesgosPlanContingenciaDetalleReversa"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDAccionContingencia") = Me.IDAccionContingencia
            .Fields("IDEdicion") = Me.IDEdicion
            If IsDate(Me.FechaInicio) Then
                .Fields("FechaInicio") = Me.FechaInicio
            End If
            If IsDate(Me.FechaFinPrevista) Then
                .Fields("FechaFinPrevista") = Me.FechaFinPrevista
            End If
            If IsDate(Me.FechaFinReal) Then
                .Fields("FechaFinReal") = Me.FechaFinReal
            End If
        .Update
        
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Eliminar( _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    If Me.IDAccionContingencia = "" Then
        p_Error = "Falta IDAccionContingencia"
        Err.Raise 1000
    End If
    m_SQL = "DELETE * " & _
            "FROM TbRiesgosPlanContingenciaDetalleReversa " & _
            "WHERE IDAccionContingencia=" & Me.IDAccionContingencia & ";"
    
    getdb().Execute m_SQL
   
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Eliminar ha devuelto el error: " & Err.Description
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
        .Add "IDAccionContingencia"
        .Add "IDEdicion"
        .Add "FechaInicio"
        .Add "FechaFinPrevista"
        .Add "FechaFinReal"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PCAccionReversa.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDAccionContingencia" Then
        getPropiedad = Me.IDAccionContingencia
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        getPropiedad = Me.IDEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicio" Then
        getPropiedad = Me.FechaInicio
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinPrevista" Then
        getPropiedad = Me.FechaFinPrevista
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinReal" Then
        getPropiedad = Me.FechaFinReal
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en PCAccionReversa.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PCAccionReversa.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDAccionContingencia" Then
        Me.IDAccionContingencia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        Me.IDEdicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicio" Then
        Me.FechaInicio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinPrevista" Then
        Me.FechaFinPrevista = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinReal" Then
        Me.FechaFinReal = m_ValorPropiedad
        Exit Function
    End If
    
    
    p_Error = "Propiedad no reconocida en PCAccionReversa.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PCAccionReversa.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


