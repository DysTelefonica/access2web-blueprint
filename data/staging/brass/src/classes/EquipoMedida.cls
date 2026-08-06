VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EquipoMedida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDEquipoMedida As String
Public Nombre As String
Public NS As String
Public PN As String
Public FechaInicioServicio As String
Public FechaFinServicio As String
Public Marca As String
Public Modelo As String
Public Descripcion As String

Public ParaAlta As EnumSino
Private m_sNombreParaCombo As String
Private m_sIDEquipoMedidaCalculado As String
Private m_objCalibraciones As Scripting.Dictionary
Private m_objUltimaCalibracion As EquipoMedidaCalibracion
Private m_sEstadoCalibracion As String
Private m_objEventos As Scripting.Dictionary
Private m_eCalibrado As EnumSino

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get UltimaCalibracion() As EquipoMedidaCalibracion
    
     On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objUltimaCalibracion Is Nothing Then
        Set UltimaCalibracion = m_objUltimaCalibracion
        Exit Property
    End If
    Set m_objUltimaCalibracion = Constructor.getUltimaCalibracion(Me.IDEquipoMedida, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UltimaCalibracion = m_objUltimaCalibracion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EventoEquipoMedida.UltimaCalibracion ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set UltimaCalibracion(ByVal objNewValue As EquipoMedidaCalibracion)

    Set m_objUltimaCalibracion = objNewValue

End Property

Public Property Get EstadoCalibracion() As String
    
     On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_sEstadoCalibracion <> "" Then
        EstadoCalibracion = m_sEstadoCalibracion
        Exit Property
    End If
    If IsDate(Me.FechaFinServicio) Then
        m_sEstadoCalibracion = "Fuera de servicio"
        EstadoCalibracion = m_sEstadoCalibracion
        Exit Property
    End If
    If Me.UltimaCalibracion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_sEstadoCalibracion = "Sin Calibraciones"
    Else
        If CDate(Me.UltimaCalibracion.FechaFinCalibracion) >= Date Then
            m_sEstadoCalibracion = "Calibrado"
        Else
            m_sEstadoCalibracion = "Fuera de Calibración"
        End If
    End If
    
    EstadoCalibracion = m_sEstadoCalibracion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EventoEquipoMedida.EstadoCalibracion ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property



Public Property Get Calibrado() As EnumSino
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Not m_eCalibrado = Empty Then
        Calibrado = m_eCalibrado
        Exit Property
    End If
    If Me.EstadoCalibracion = "Calibrado" Then
        m_eCalibrado = EnumSino.Sí
    Else
        m_eCalibrado = EnumSino.No
    End If
    
    Calibrado = m_eCalibrado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EquipoMedida.Calibrado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Eventos() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Not m_objEventos Is Nothing Then
        Set Eventos = m_objEventos
        Exit Property
    End If
    If Me.IDEquipoMedida = "" Then
        Exit Property
    End If
    Set m_objEventos = Constructor.getEventosConEquipoMedida(Me.IDEquipoMedida, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Eventos = m_objEventos
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EquipoMedida.Eventos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get NombreParaCombo() As String
     
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sNombreParaCombo <> "" Then
        NombreParaCombo = m_sNombreParaCombo
        Exit Property
    End If
    m_sNombreParaCombo = Replace(Me.Nombre, ";", ":") & " (" & Replace(Me.NS, ";", ":") & " )"
    NombreParaCombo = m_sNombreParaCombo
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EquipoMedida.NombreParaCombo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get IDEquipoMedidaCalculado() As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    m_sIDEquipoMedidaCalculado = DameID("TbEquiposMedida", "IDEquipoMedida")
    
    IDEquipoMedidaCalculado = m_sIDEquipoMedidaCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EquipoMedida.IDEquipoMedidaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get Calibraciones() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objCalibraciones Is Nothing Then
        Set Calibraciones = m_objCalibraciones
        Exit Property
    End If
    If Me.IDEquipoMedida = "" Then
        Exit Property
        
    End If
    Set m_objCalibraciones = Constructor.getEquipoMedidaCalibraciones(Me.IDEquipoMedida, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Calibraciones = m_objCalibraciones
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EquipoMedida.Calibraciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property
Public Property Set Calibraciones(NuevoValor As Variant)

    Set m_objCalibraciones = NuevoValor

End Property



Private Function getEquipoMedidaEnActividad( _
                                            Optional p_IDEventoEquipoMedida As String, _
                                            Optional p_IDEquipoMedida As String, _
                                            Optional p_IDEvento As String, _
                                            Optional ByRef p_Error As String _
                                            ) As EventoEquipoMedida

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    
    On Error GoTo errores
    If p_IDEventoEquipoMedida = "" Or (p_IDEquipoMedida = "" And p_IDEvento = "") Then
        Exit Function
    End If
    If p_IDEventoEquipoMedida <> "" Then
        m_SQL = "SELECT * FROM TbEventosEquipoMedida " & _
                "WHERE IDEventoEquipoMedida=" & p_IDEventoEquipoMedida & ";"
    Else
        m_SQL = "SELECT * FROM TbEventosEquipoMedida " & _
                "WHERE IDEvento='" & p_IDEvento & _
                "' AND IDEquipoMedida=" & p_IDEquipoMedida & ";"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getEquipoMedidaEnActividad = New EventoEquipoMedida
        For Each m_Campo In getEquipoMedidaEnActividad.ColCampos
            getEquipoMedidaEnActividad.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getEquipoMedidaEnActividad ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Registrar( _
                            Optional ByRef p_EquipoAlInicio As EquipoMedida, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_NSRepetido As EnumSino
    Dim m_ID As String
    
    On Error GoTo errores
    If Me.ParaAlta = Empty Then
        p_Error = "No se sabe si es para alta o no"
        Err.Raise 1000
    End If
    If Me.ParaAlta = EnumSino.No Then
        If p_EquipoAlInicio Is Nothing Then
            Set p_EquipoAlInicio = Constructor.getEquipoMedida(Me.IDEquipoMedida, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
        If p_EquipoAlInicio Is Nothing Then
            p_Error = "Es en modo edición y no se encuentra el Estado inicial del Equipo de medida"
            Err.Raise 1000
        End If
    End If
    With Me
        If Me.Nombre = "" Then
            p_Error = "Falta el nombre"
            Err.Raise 1000
        End If
        If Me.NS = "" Then
            p_Error = "Falta el NS"
            Err.Raise 1000
        End If
        If Me.PN = "" Then
            p_Error = "Falta el PN"
            Err.Raise 1000
        End If
        If Me.FechaInicioServicio = "" Then
            p_Error = "Falta el FechaInicioServicio"
            Err.Raise 1000
        End If
        If Me.ParaAlta = EnumSino.No Then
            If p_EquipoAlInicio.NS <> Me.NS Then
                m_NSRepetido = NSRepetido(p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If m_NSRepetido = EnumSino.Sí Then
                    p_Error = "El Número de Serie ya está siendo usado por otro equipo de medida"
                    Err.Raise 1000
                End If
            End If
        End If
        
        If Me.FechaFinServicio <> "" Then
            If CDate(Me.FechaFinServicio) < CDate(Me.FechaInicioServicio) Then
                p_Error = "La fecha fin de servicio no puede ser anterior a la de comienzo"
                Err.Raise 1000
            End If
        End If
        If Me.IDEquipoMedida = "" Then
           
            m_ID = Me.IDEquipoMedidaCalculado
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Me.IDEquipoMedida = m_ID
        
        End If
        
    End With
    m_SQL = "SELECT * " & _
                "FROM TbEquiposMedida " & _
                "WHERE IDEquipoMedida=" & Me.IDEquipoMedida & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                .Fields("IDEquipoMedida") = Me.IDEquipoMedida
                If Me.Descripcion <> "" Then
                    .Fields("Descripcion") = Me.Descripcion
                End If
                If Me.FechaFinServicio <> "" Then
                    .Fields("FechaFinServicio") = Me.FechaFinServicio
                End If
                If Me.FechaInicioServicio <> "" Then
                    .Fields("FechaInicioServicio") = Me.FechaInicioServicio
                End If
                If Me.Marca <> "" Then
                    .Fields("Marca") = Me.Marca
                End If
                If Me.Modelo <> "" Then
                    .Fields("Modelo") = Me.Modelo
                End If
                If Me.Nombre <> "" Then
                    .Fields("Nombre") = Me.Nombre
                End If
                If Me.NS <> "" Then
                    .Fields("NS") = Me.NS
                End If
                If Me.PN <> "" Then
                    .Fields("PN") = Me.PN
                End If
                
            .Update
        Else
            .Edit
                If Me.Descripcion <> "" Then
                    .Fields("Descripcion") = Me.Descripcion
                Else
                    .Fields("Descripcion") = Null
                End If
                If Me.FechaFinServicio <> "" Then
                    .Fields("FechaFinServicio") = Me.FechaFinServicio
                Else
                    .Fields("FechaFinServicio") = Null
                End If
                If Me.FechaInicioServicio <> "" Then
                    .Fields("FechaInicioServicio") = Me.FechaInicioServicio
                Else
                    .Fields("FechaInicioServicio") = Null
                End If
                If Me.Marca <> "" Then
                    .Fields("Marca") = Me.Marca
                
                Else
                    .Fields("Marca") = Null
                End If
                If Me.Modelo <> "" Then
                    .Fields("Modelo") = Me.Modelo
                Else
                    .Fields("Modelo") = Null
                End If
                If Me.Nombre <> "" Then
                    .Fields("Nombre") = Me.Nombre
                Else
                    .Fields("Nombre") = Null
                End If
                If Me.NS <> "" Then
                    .Fields("NS") = Me.NS
                Else
                    .Fields("NS") = Null
                End If
                If Me.PN <> "" Then
                    .Fields("PN") = Me.PN
                Else
                    .Fields("PN") = Null
                End If
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Me.ParaAlta = EnumSino.No Then
        If Me.FechaFinServicio <> p_EquipoAlInicio.FechaFinServicio Then
            Set m_ObjEntorno.ColEquiposMedidaActivos = Constructor.getEquiposMedidaActivos(p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EquipoMedida.Registrar ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function Eliminar(Optional ByRef p_Error As String) As String
    
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If Me.Eventos Is Nothing Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "El aparato de medida está registrado en al menos un evento"
        Err.Raise 1000
    End If
    
    
    m_SQL = "Delete * " & _
                "FROM TbEquiposMedida " & _
                "WHERE IDEquipoMedida=" & Me.IDEquipoMedida & ";"
    getdb().Execute m_SQL
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EquipoMedida.Eliminar ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Private Function NSRepetido(Optional ByRef p_Error As String) As EnumSino
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
   
    On Error GoTo errores
    With Me
        If Me.NS = "" Then
            p_Error = "Falta el NS"
            Err.Raise 1000
        End If
    End With
    m_SQL = "SELECT NS " & _
            "FROM TbEquiposMedida " & _
            "WHERE NS='" & Me.NS & "';"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            NSRepetido = EnumSino.No
        Else
            NSRepetido = EnumSino.Sí
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EquipoMedida.EquipoRepetido ha devuelto el error:" & vbNewLine & Err.Description
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
        .Add "IDEquipoMedida"
        .Add "Nombre"
        .Add "NS"
        .Add "PN"
        .Add "FechaInicioServicio"
        .Add "FechaFinServicio"
        .Add "Marca"
        .Add "Modelo"
        .Add "Descripcion"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EquipoMedida.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDEquipoMedida" Then
        getPropiedad = Me.IDEquipoMedida
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        getPropiedad = Me.Nombre
        Exit Function
    End If
    If m_NombrePropiedad = "NS" Then
        getPropiedad = Me.NS
        Exit Function
    End If
    If m_NombrePropiedad = "PN" Then
        getPropiedad = Me.PN
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicioServicio" Then
        getPropiedad = Me.FechaInicioServicio
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinServicio" Then
        getPropiedad = Me.FechaFinServicio
        Exit Function
    End If
    If m_NombrePropiedad = "Marca" Then
        getPropiedad = Me.Marca
        Exit Function
    End If
    If m_NombrePropiedad = "Modelo" Then
        getPropiedad = Me.Modelo
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en EquipoMedida.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EquipoMedida.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDEquipoMedida" Then
        Me.IDEquipoMedida = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        Me.Nombre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NS" Then
        Me.NS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PN" Then
        Me.PN = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicioServicio" Then
        Me.FechaInicioServicio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinServicio" Then
        Me.FechaFinServicio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Marca" Then
        Me.Marca = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Modelo" Then
        Me.Modelo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en EquipoMedida.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EquipoMedida.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function



