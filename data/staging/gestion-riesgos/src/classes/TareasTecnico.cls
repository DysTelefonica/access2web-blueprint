VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TareasTecnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_objColEdicionesAPuntoDeCaducarSinPropuesta As Scripting.Dictionary
Private m_objColEdicionesPropuestasRechazadas As Scripting.Dictionary
Private m_objColRiesgosAceptadosRechazados As Scripting.Dictionary
Private m_objColRiesgosRetiradosRechazados As Scripting.Dictionary

Private m_sNEdicionesAPuntoDeCaducarSinPropuesta As String
Private m_sNEdicionesPropuestasRechazadas As String
Private m_sNRiesgosAceptadosRechazados As String
Private m_sNRiesgosRetiradosRechazados As String

Private m_eHayTareasPendientes As EnumSiNo
Private m_sTareasTotales As String

Public Error As String
Private m_Error As String

Public Property Get ColEdicionesAPuntoDeCaducarSinPropuesta() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEdicionesAPuntoDeCaducarSinPropuesta Is Nothing Then
        Set ColEdicionesAPuntoDeCaducarSinPropuesta = m_objColEdicionesAPuntoDeCaducarSinPropuesta
        Exit Property
    End If
    Set m_objColEdicionesAPuntoDeCaducarSinPropuesta = getEdicionesAPuntoDeCaducarSinPropuesta(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColEdicionesAPuntoDeCaducarSinPropuesta = m_objColEdicionesAPuntoDeCaducarSinPropuesta
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TareasTecnico.ColEdicionesAPuntoDeCaducarSinPropuesta ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property
Public Property Set ColEdicionesAPuntoDeCaducarSinPropuesta(p_Valor As Variant)

    Set m_objColEdicionesAPuntoDeCaducarSinPropuesta = p_Valor
End Property


Public Property Get ColEdicionesPropuestasRechazadas() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColEdicionesPropuestasRechazadas Is Nothing Then
        Set ColEdicionesPropuestasRechazadas = m_objColEdicionesPropuestasRechazadas
        Exit Property
    End If
    Set m_objColEdicionesPropuestasRechazadas = getEdicionesPropuestasRechazadas(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColEdicionesPropuestasRechazadas = m_objColEdicionesPropuestasRechazadas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TareasTecnico.ColEdicionesPropuestasRechazadas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property
Public Property Set ColEdicionesPropuestasRechazadas(p_Valor As Variant)

    Set m_objColEdicionesPropuestasRechazadas = p_Valor
End Property


Public Property Get ColRiesgosAceptadosRechazados() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosAceptadosRechazados Is Nothing Then
        Set ColRiesgosAceptadosRechazados = m_objColRiesgosAceptadosRechazados
        Exit Property
    End If
    Set m_objColRiesgosAceptadosRechazados = getRiesgosAceptadosRechazados(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColRiesgosAceptadosRechazados = m_objColRiesgosAceptadosRechazados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TareasTecnico.ColRiesgosAceptadosRechazados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosAceptadosRechazados(ByVal objNewValue As Variant)

    Set m_objColRiesgosAceptadosRechazados = objNewValue

End Property

Public Property Get ColRiesgosRetiradosRechazados() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosRetiradosRechazados Is Nothing Then
        Set ColRiesgosRetiradosRechazados = m_objColRiesgosRetiradosRechazados
        Exit Property
    End If
    Set m_objColRiesgosRetiradosRechazados = getRiesgosRetiradosRechazados(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColRiesgosRetiradosRechazados = m_objColRiesgosRetiradosRechazados
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TareasTecnico.ColRiesgosRetiradosRechazados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosRetiradosRechazados(ByVal objNewValue As Variant)
    
    Set m_objColRiesgosRetiradosRechazados = objNewValue

End Property

Public Property Get NRiesgosAceptadosRechazados() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.ColRiesgosAceptadosRechazados Is Nothing Then
        m_sNRiesgosAceptadosRechazados = "0"
    Else
        m_sNRiesgosAceptadosRechazados = CStr(Me.ColRiesgosAceptadosRechazados.Count)
    End If
    NRiesgosAceptadosRechazados = m_sNRiesgosAceptadosRechazados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TareasTecnico.NRiesgosAceptadosRechazados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get NRiesgosRetiradosRechazados() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.ColRiesgosRetiradosRechazados Is Nothing Then
        m_sNRiesgosRetiradosRechazados = "0"
    Else
        m_sNRiesgosRetiradosRechazados = CStr(Me.ColRiesgosRetiradosRechazados.Count)
    End If
    NRiesgosRetiradosRechazados = m_sNRiesgosRetiradosRechazados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TareasTecnico.NRiesgosRetiradosRechazados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property




Public Property Get NEdicionesAPuntoDeCaducarSinPropuesta() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ColEdicionesAPuntoDeCaducarSinPropuesta Is Nothing Then
        m_sNEdicionesAPuntoDeCaducarSinPropuesta = "0"
    Else
        m_sNEdicionesAPuntoDeCaducarSinPropuesta = CStr(Me.ColEdicionesAPuntoDeCaducarSinPropuesta.Count)
    End If
    NEdicionesAPuntoDeCaducarSinPropuesta = m_sNEdicionesAPuntoDeCaducarSinPropuesta
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TareasTecnico.NEdicionesAPuntoDeCaducarSinPropuesta ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NEdicionesPropuestasRechazadas() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ColEdicionesPropuestasRechazadas Is Nothing Then
        m_sNEdicionesPropuestasRechazadas = "0"
    Else
        m_sNEdicionesPropuestasRechazadas = CStr(Me.ColEdicionesPropuestasRechazadas.Count)
    End If
    NEdicionesPropuestasRechazadas = m_sNEdicionesPropuestasRechazadas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TareasTecnico.NEdicionesPropuestasRechazadas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TareasTotales() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sTareasTotales = (CInt(Me.NEdicionesAPuntoDeCaducarSinPropuesta) + CInt(Me.NRiesgosAceptadosRechazados) + _
                                CInt(Me.NRiesgosRetiradosRechazados) + CInt(Me.NEdicionesPropuestasRechazadas))
   
    TareasTotales = m_sTareasTotales
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TareasTecnico.TareasTotales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let TareasTotales(ByVal sNewValue As String)

    m_sTareasTotales = sNewValue

End Property



Public Property Get HayTareasPendientes() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not IsNumeric(Me.TareasTotales) Then
        Exit Property
    End If
    If CInt(Me.TareasTotales) = 0 Then
        m_eHayTareasPendientes = EnumSiNo.No
    Else
        m_eHayTareasPendientes = EnumSiNo.Sí
    End If
    HayTareasPendientes = m_eHayTareasPendientes
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TareasTecnico.HayTareasPendientes ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Let HayTareasPendientes(valor As EnumSiNo)

    m_eHayTareasPendientes = valor
    
End Property

Private Function getEdicionesAPuntoDeCaducarSinPropuesta( _
                                                            Optional ByRef p_Error As String _
                                                            ) As Scripting.Dictionary
                    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Edicion As Edicion
        
    On Error GoTo errores
    
    m_SQL = "SELECT TbProyectosEdiciones.* " & _
                "FROM TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
                "WHERE (((TbProyectos.FechaCierre) Is Null) " & _
                "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                "AND ((TbProyectosEdiciones.FechaPreparadaParaPublicar) Is Null) " & _
                "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null) " & _
                "AND ((DateDiff('d',Now(),[TbProyectosEdiciones].[FechaMaxProximaPublicacion]))<=" & _
                        m_ObjEntorno.JPDiasPreviosParaElAviso & ") " & _
                "AND ((TbProyectos.CadenaNombreAutorizados) Like'*" & m_ObjUsuarioConectado.Nombre & "*'));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Edicion = New Edicion
            For Each m_Campo In m_Edicion.ColCampos
                m_Edicion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getEdicionesAPuntoDeCaducarSinPropuesta Is Nothing Then
                Set getEdicionesAPuntoDeCaducarSinPropuesta = New Scripting.Dictionary
                getEdicionesAPuntoDeCaducarSinPropuesta.CompareMode = TextCompare
            End If
            If Not getEdicionesAPuntoDeCaducarSinPropuesta.Exists(CStr(m_Edicion.IDEdicion)) Then
                getEdicionesAPuntoDeCaducarSinPropuesta.Add CStr(m_Edicion.IDEdicion), m_Edicion
            End If
            Set m_Edicion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getEdicionesAPuntoDeCaducarSinPropuesta ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getEdicionesPropuestasRechazadas( _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
                    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Edicion As Edicion
        
    On Error GoTo errores
    
    m_SQL = "SELECT TbProyectosEdiciones.* " & _
                "FROM TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
                "WHERE (((TbProyectos.FechaCierre) Is Null) " & _
                "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                "AND ((TbProyectosEdiciones.FechaPreparadaParaPublicar) Is Null) " & _
                "AND (Not(TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null) " & _
                "AND ((TbProyectos.CadenaNombreAutorizados) Like'*" & m_ObjUsuarioConectado.Nombre & "*'));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Edicion = New Edicion
            For Each m_Campo In m_Edicion.ColCampos
                m_Edicion.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getEdicionesPropuestasRechazadas Is Nothing Then
                Set getEdicionesPropuestasRechazadas = New Scripting.Dictionary
                getEdicionesPropuestasRechazadas.CompareMode = TextCompare
            End If
            If Not getEdicionesPropuestasRechazadas.Exists(CStr(m_Edicion.IDEdicion)) Then
                getEdicionesPropuestasRechazadas.Add CStr(m_Edicion.IDEdicion), m_Edicion
            End If
            Set m_Edicion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getEdicionesPropuestasRechazadas ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getRiesgosAceptadosRechazados( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Riesgo As riesgo
        
    On Error GoTo errores
    
    m_SQL = "SELECT TbRiesgos.* " & _
            "FROM (TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
            "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
            "WHERE ((Not (TbRiesgos.FechaJustificacionAceptacionRiesgo) Is Null) " & _
            "AND ((TbProyectos.FechaCierre) Is Null) " & _
            "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) AND ((TbRiesgos.Mitigacion)='Aceptar') " & _
            "AND (Not(TbRiesgos.FechaRechazoAceptacionPorCalidad) Is Null) " & _
            "AND ((TbProyectos.CadenaNombreAutorizados) Like'*" & m_ObjUsuarioConectado.Nombre & "*'));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Riesgo = New riesgo
            For Each m_Campo In m_Riesgo.ColCampos
                m_Riesgo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getRiesgosAceptadosRechazados Is Nothing Then
                Set getRiesgosAceptadosRechazados = New Scripting.Dictionary
                getRiesgosAceptadosRechazados.CompareMode = TextCompare
            End If
            If Not getRiesgosAceptadosRechazados.Exists(CStr(m_Riesgo.IDRiesgo)) Then
                getRiesgosAceptadosRechazados.Add CStr(m_Riesgo.IDRiesgo), m_Riesgo
            End If
            Set m_Riesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosAceptadosRechazados ha devuelto el error: " & Err.Description
    End If
End Function


Private Function getRiesgosRetiradosRechazados( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Riesgo As riesgo
        
    On Error GoTo errores
    
    m_SQL = "SELECT TbRiesgos.* " & _
            "FROM (TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
            "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
            "WHERE ((Not (TbRiesgos.FechaJustificacionRetiroRiesgo) Is Null) " & _
            "AND ((TbProyectos.FechaCierre) Is Null) " & _
            "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
            "AND (Not(TbRiesgos.FechaRechazoRetiroPorCalidad) Is Null) " & _
            "AND ((TbProyectos.CadenaNombreAutorizados) Like'*" & m_ObjUsuarioConectado.Nombre & "*'));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Riesgo = New riesgo
            For Each m_Campo In m_Riesgo.ColCampos
                m_Riesgo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getRiesgosRetiradosRechazados Is Nothing Then
                Set getRiesgosRetiradosRechazados = New Scripting.Dictionary
                getRiesgosRetiradosRechazados.CompareMode = TextCompare
            End If
            If Not getRiesgosRetiradosRechazados.Exists(CStr(m_Riesgo.IDRiesgo)) Then
                getRiesgosRetiradosRechazados.Add CStr(m_Riesgo.IDRiesgo), m_Riesgo
            End If
            Set m_Riesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosRetiradosRechazados ha devuelto el error: " & Err.Description
    End If
End Function





















