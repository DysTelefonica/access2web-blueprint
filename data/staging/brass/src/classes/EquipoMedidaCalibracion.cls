VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EquipoMedidaCalibracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDCalibracion As String
Public IDEquipoMedida As String
Public FechaCalibracion As String
Public FechaFinCalibracion As String
Public EmpresaCalibradora As String
Public Observaciones As String

Public ParaAlta As EnumSino
Private m_sIDCalibracionCalculado As String
Private m_objEquipoMedida As EquipoMedida
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get IDCalibracionCalculado() As String

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    m_sIDCalibracionCalculado = DameID("TbEquiposMedidaCalibraciones", "IDCalibracion")
    
    IDCalibracionCalculado = m_sIDCalibracionCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EquipoMedidaCalibracion.IDCalibracionCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get EquipoMedida() As EquipoMedida

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Not m_objEquipoMedida Is Nothing Then
        Set EquipoMedida = m_objEquipoMedida
        Exit Property
    End If
    If Me.IDEquipoMedida = "" Then
        Exit Property
    End If
    Set m_objEquipoMedida = Constructor.getEquipoMedida(Me.IDEquipoMedida, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EquipoMedida = m_objEquipoMedida
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EquipoMedidaCalibracion.EquipoMedida ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Function Registrar( _
                            Optional ByRef p_CalibracionAlInicio As EquipoMedidaCalibracion, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If Me.ParaAlta = Empty Then
        p_Error = "No se sabe si es para alta o no"
        Err.Raise 1000
    End If
    If Me.ParaAlta = EnumSino.No Then
        If p_CalibracionAlInicio Is Nothing Then
            Set p_CalibracionAlInicio = Constructor.getEquipoMedidaCalibracion(Me.IDCalibracion, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
        If p_CalibracionAlInicio Is Nothing Then
            p_Error = "Es en modo edición y no se encuentra el Estado inicial de la calibración"
            Err.Raise 1000
        End If
    Else
        Me.IDCalibracion = Me.IDCalibracionCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If Not IsDate(Me.FechaCalibracion) Then
        p_Error = "La fecha inicial es obligatoria"
        Err.Raise 1000
    End If
    If Not IsDate(Me.FechaFinCalibracion) Then
        p_Error = "La fecha Final es obligatoria"
        Err.Raise 1000
    End If
    If CDate(Me.FechaFinCalibracion) < CDate(Me.FechaCalibracion) Then
        p_Error = "La fecha Final ha de ser posterior a la inicial"
        Err.Raise 1000
    End If
    Me.EmpresaCalibradora = "LA ARMADA"
    m_SQL = "SELECT * " & _
                "FROM TbEquiposMedidaCalibraciones " & _
                "WHERE IDCalibracion=" & Me.IDCalibracion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Me.ParaAlta = EnumSino.No Then
            If .EOF Then
                p_Error = "No es para alta y no se encuentra la calibración registrada en el Sistema"
                Err.Raise 1000
            End If
        End If
        If .EOF Then
            .AddNew
                .Fields("IDCalibracion") = Me.IDCalibracion
                .Fields("IDEquipoMedida") = Me.IDEquipoMedida
                .Fields("FechaCalibracion") = Me.FechaCalibracion
                .Fields("FechaFinCalibracion") = Me.FechaFinCalibracion
                .Fields("EmpresaCalibradora") = Me.EmpresaCalibradora
                If Me.Observaciones <> "" Then
                    .Fields("Observaciones") = Me.Observaciones
                End If
            .Update
        Else
            .Edit
                .Fields("IDEquipoMedida") = Me.IDEquipoMedida
                .Fields("FechaCalibracion") = Me.FechaCalibracion
                .Fields("FechaFinCalibracion") = Me.FechaFinCalibracion
                .Fields("EmpresaCalibradora") = Me.EmpresaCalibradora
                If Me.Observaciones <> "" Then
                    .Fields("Observaciones") = Me.Observaciones
                Else
                    .Fields("Observaciones") = Null
                End If
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EquipoMedidaCalibracion.Registrar ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function Eliminar(Optional ByRef p_Error As String) As String
    
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If Me.IDCalibracion = "" Then
        p_Error = "No se conoce la IDCalibración"
        Err.Raise 1000
    End If
    m_SQL = "Delete * " & _
                "FROM TbEquiposMedidaCalibraciones " & _
                "WHERE IDCalibracion=" & Me.IDCalibracion & ";"
    getdb().Execute m_SQL
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EquipoMedidaCalibracion.Eliminar ha devuelto el error:" & vbNewLine & Err.Description
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
        .Add "IDCalibracion"
        .Add "IDEquipoMedida"
        .Add "FechaCalibracion"
        .Add "FechaFinCalibracion"
        .Add "EmpresaCalibradora"
        .Add "Observaciones"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EquipoMedidaCalibracion.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDCalibracion" Then
        getPropiedad = Me.IDCalibracion
        Exit Function
    End If
    If m_NombrePropiedad = "IDEquipoMedida" Then
        getPropiedad = Me.IDEquipoMedida
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCalibracion" Then
        getPropiedad = Me.FechaCalibracion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinCalibracion" Then
        getPropiedad = Me.FechaFinCalibracion
        Exit Function
    End If
    If m_NombrePropiedad = "EmpresaCalibradora" Then
        getPropiedad = Me.EmpresaCalibradora
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en EquipoMedidaCalibracion.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EquipoMedidaCalibracion.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDCalibracion" Then
        Me.IDCalibracion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEquipoMedida" Then
        Me.IDEquipoMedida = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCalibracion" Then
        Me.FechaCalibracion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinCalibracion" Then
        Me.FechaFinCalibracion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EmpresaCalibradora" Then
        Me.EmpresaCalibradora = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en EquipoMedidaCalibracion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EquipoMedidaCalibracion.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


