Attribute VB_Name = "Instalacion"
Option Compare Database
Option Explicit

Public Function instalar(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    RegularizarRiesgosMaterializados p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    RegularizarTiemposAceptadosRetirados p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    RegularizarCodigoCompletoRiesgos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    RegularizarOrigenesVacios p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    RellenarFechaCierreProyecto p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarCadenaAutorizados p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarNombreIconoRiesgosActivos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarCampoEstadoEnActivos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegularizarCambiosCompletosEnProyectos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarSiHayErroresYNombreNodosGral p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método instalar ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function RellenarFechaCierreProyecto(Optional ByRef p_Error As String) As String
    Dim m_SQL As String
    
    On Error GoTo errores
    m_SQL = "UPDATE TbProyectos SET TbProyectos.FechaCierre = [FechaCierre];"
    getdb().Execute m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarFechaCierreProyecto ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function RellenarCadenaAutorizados(Optional ByRef p_Error As String) As String
    
    Dim m_ID As Variant
    Dim m_Proyecto As Proyecto
    Dim m_SQL As String
    Dim m_CadenaCalculada As String
    
    On Error GoTo errores
    For Each m_ID In m_ObjEntorno.ColProyectosTotales
        Set m_Proyecto = m_ObjEntorno.ColProyectosTotales(m_ID)
        m_CadenaCalculada = m_Proyecto.CadenaNombreAutorizadosCalculados
        If m_CadenaCalculada = "" Then
            Debug.Print m_Proyecto.IDProyecto, "SIN CADENA"
        End If
        If m_Proyecto.CadenaNombreAutorizados <> m_CadenaCalculada Then
            m_SQL = "UPDATE TbProyectos SET CadenaNombreAutorizados ='" & m_CadenaCalculada & "' " & _
             "WHERE IDProyecto=" & m_ID & ";"
            getdb().Execute m_SQL
            VBA.DoEvents
            Debug.Print m_Proyecto.IDProyecto, m_Proyecto.CadenaNombreAutorizados, m_CadenaCalculada
            VBA.DoEvents
        Else
            VBA.DoEvents
            Debug.Print m_Proyecto.IDProyecto, "IGUAL"
            VBA.DoEvents
        End If
siguiente:
        Set m_Proyecto = Nothing
    Next
    
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarCadenaAutorizados ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function RellenarNombreParaNodo(Optional ByRef p_Error As String) As String
    
    Dim m_ID As Variant
    Dim m_Proyecto As Proyecto
    Dim m_SQL As String
    Dim m_NombreParaNodoCalculado As String
    
    On Error GoTo errores
    For Each m_ID In m_ObjEntorno.ColProyectosTotales
       ' If CStr(m_ID) = "98" Then Stop
        Set m_Proyecto = m_ObjEntorno.ColProyectosTotales(m_ID)
        m_NombreParaNodoCalculado = m_Proyecto.NombreParaNodoCalculado
        If m_NombreParaNodoCalculado = "" Then
            Debug.Print m_Proyecto.IDProyecto, "SIN NombreParaNodoCalculado"
        End If
        If m_Proyecto.NombreParaNodo <> m_NombreParaNodoCalculado Then
            If m_NombreParaNodoCalculado <> "" Then
                 m_SQL = "UPDATE TbProyectos SET NombreParaNodo ='" & m_NombreParaNodoCalculado & "' " & _
                        "WHERE IDProyecto=" & m_ID & ";"
            Else
                m_SQL = "UPDATE TbProyectos SET NombreParaNodo =Null " & _
                        "WHERE IDProyecto=" & m_ID & ";"
            End If
           
            getdb().Execute m_SQL
            VBA.DoEvents
            Debug.Print m_Proyecto.IDProyecto, m_Proyecto.NombreParaNodo, m_NombreParaNodoCalculado
            VBA.DoEvents
        Else
            VBA.DoEvents
            Debug.Print m_Proyecto.IDProyecto, "IGUAL"
            VBA.DoEvents
        End If
siguiente:
        Set m_Proyecto = Nothing
    Next
    
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarNombreParaNodo ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function RellenarPriorizacionAceptadosNoActivos(Optional ByRef p_Error As String) As String
    
    Dim m_Proyecto As Proyecto
    Dim m_IDProyecto As Variant
    Dim m_Edicion As Edicion
    Dim m_IDEdicion As Variant
    Dim m_ColRiesgos As Scripting.Dictionary
    Dim m_IdRiesgo As Variant
    Dim m_Riesgo As riesgo
    Dim i As Long
    
    Dim m_SQL As String
    Dim m_CadenaCalculada As String
    
    On Error GoTo errores
    For Each m_IDProyecto In m_ObjEntorno.ColProyectosTotales
        Set m_Proyecto = m_ObjEntorno.ColProyectosTotales(m_IDProyecto)
        For Each m_IDEdicion In m_Proyecto.colEdiciones
            Set m_Edicion = m_Proyecto.colEdiciones(m_IDEdicion)
                If m_Edicion.EsActivo = EnumSiNo.No Then
                    Set m_ColRiesgos = m_Edicion.colRiesgos
                    If Not m_ColRiesgos Is Nothing Then
                        For Each m_IdRiesgo In m_ColRiesgos
                            Set m_Riesgo = m_ColRiesgos(m_IdRiesgo)
                            
                            
                        Next
                    End If
                    Set m_ColRiesgos = Nothing
                End If
            Set m_Edicion = Nothing
        Next
        
'        m_CadenaCalculada = m_Proyecto.CadenaNombreAutorizadosCalculados
'        If m_CadenaCalculada = "" Then
'            Debug.Print m_Proyecto.IDProyecto, "SIN CADENA"
'        End If
'        If m_Proyecto.CadenaNombreAutorizados <> m_CadenaCalculada Then
'            m_SQL = "UPDATE TbProyectos SET CadenaNombreAutorizados ='" & m_CadenaCalculada & "' " & _
'             "WHERE IDProyecto=" & m_ID & ";"
'            getdb().Execute m_SQL
'            VBA.DoEvents
'            Debug.Print m_Proyecto.IDProyecto, m_Proyecto.CadenaNombreAutorizados, m_CadenaCalculada
'            VBA.DoEvents
'        Else
'            VBA.DoEvents
'            Debug.Print m_Proyecto.IDProyecto, "IGUAL"
'            VBA.DoEvents
'        End If
siguiente:
        Set m_Proyecto = Nothing
    Next
    
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarPriorizacionAceptadosNoActivos ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function RepasarFechasMaxProxPublicacion(Optional ByRef p_Error As String) As String
    
   
    Dim m_Edicion As Edicion
    Dim m_Proyecto As Proyecto
    Dim m_FechaUltimaPublicacion As String
    Dim m_FechaMaxEdicion As String
    Dim m_FechaMaxEdicionCalculado As String
    Dim m_FechaMaxProyecto As String
    Dim m_UltimaEdicionPublicada As Edicion
    Dim m_IDEdicion As Variant
    Dim m_ColEdiciones As Scripting.Dictionary
    Dim m_Expediente As Expediente
    Dim m_SQL As String
    
    On Error GoTo errores
    Set m_ColEdiciones = Constructor.getEdicionesActivas(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColEdiciones Is Nothing Then
        Exit Function
    End If
    For Each m_IDEdicion In m_ColEdiciones
        m_FechaUltimaPublicacion = ""
        m_FechaMaxEdicion = ""
        m_FechaMaxProyecto = ""
        
        Set m_Edicion = m_ColEdiciones(m_IDEdicion)
        Set m_Proyecto = m_Edicion.Proyecto
        Set m_Expediente = m_Proyecto.Expediente
        'If m_Expediente.Nemotecnico = "SDR 22" Then Stop
        If Not m_Proyecto.EdicionUltimaPublicada Is Nothing Then
            If IsDate(m_Proyecto.EdicionUltimaPublicada.FechaPublicacion) Then
                m_FechaUltimaPublicacion = m_Proyecto.EdicionUltimaPublicada.FechaPublicacion
            End If
        End If
        m_FechaMaxEdicionCalculado = m_Edicion.FechaMaxProximaPublicacionCalculado
        m_FechaMaxEdicion = m_Edicion.FechaMaxProximaPublicacion
        m_SQL = "UPDATE TbProyectosEdiciones " & _
                "SET FechaMaxProximaPublicacion=#" & Format(m_FechaMaxEdicionCalculado, "mm/dd/yyyy") & "# " & _
                "WHERE IDEdicion=" & m_Edicion.IDEdicion & ";"
        getdb().Execute m_SQL
        m_SQL = "UPDATE TbProyectos " & _
                "SET FechaMaxProximaPublicacion=#" & Format(m_FechaMaxEdicionCalculado, "mm/dd/yyyy") & "# " & _
                "WHERE IDProyecto=" & m_Proyecto.IDProyecto & ";"
        getdb().Execute m_SQL
        If m_FechaMaxEdicionCalculado <> m_FechaMaxEdicion Then
            
            Debug.Print m_Proyecto.NombreProyecto & " " & "UltimaPub: " & m_FechaUltimaPublicacion & _
                        " FechaMaxEd " & m_FechaMaxEdicionCalculado
        Else
            Debug.Print m_Proyecto.NombreProyecto & " Igual"
        End If
        
       
        Set m_Edicion = Nothing
        Set m_Proyecto = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RepasarFechasMaxProxPublicacion ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function RellenarSiHayErroresYNombreNodosGral(Optional ByRef p_Error As String) As String

    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Riesgo As riesgo
    Dim m_PM As PM
    Dim m_PC As PC
    Dim m_IDPM As Variant
    Dim m_IDPC As Variant
    Dim m_ColPM As Scripting.Dictionary
    Dim m_ColPC As Scripting.Dictionary
    Dim m_PMA As PMAccion
    Dim m_PCA As PCAccion
    Dim m_IDPMA As Variant
    Dim m_IDPCA As Variant
    Dim m_ColPMA As Scripting.Dictionary
    Dim m_ColPCA As Scripting.Dictionary
    
    On Error GoTo errores
    Set m_Col = Constructor.getRiesgosActivos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Riesgo = m_Col(m_ID)
        With m_Riesgo
            VBA.DoEvents
            Debug.Print .CodigoUnicoCalculado
            VBA.DoEvents
            .GrabarNombreNodoEstado p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            .HayErrorEnRiesgoGrabar p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            .GrabarNombreNodoDesc p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            .GrabarNombreIcono p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            .GrabarColorIcono p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
        End With
        
        Set m_ColPM = m_Riesgo.ColPMs
        p_Error = m_Riesgo.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ColPM Is Nothing Then
            For Each m_IDPM In m_ColPM
                Set m_PM = m_ColPM(m_IDPM)
                With m_PM
                    .GrabarNombreIcono p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If Not .colAcciones Is Nothing Then
                        For Each m_IDPMA In .colAcciones
                            Set m_PMA = .colAcciones(m_IDPMA)
                            m_PMA.GrabarNombreIcono p_Error
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                            Set m_PMA = Nothing
                        Next
                    End If
                End With
                Set m_PM = Nothing
            Next
        End If
        Set m_ColPC = m_Riesgo.ColPCs
        p_Error = m_Riesgo.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ColPC Is Nothing Then
            For Each m_IDPC In m_ColPC
                Set m_PC = m_ColPC(m_IDPC)
                With m_PC
                    If .NombreIcono = "" Then
                        .GrabarNombreIcono p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    End If
                    If Not .colAcciones Is Nothing Then
                        For Each m_IDPCA In .colAcciones
                            Set m_PCA = .colAcciones(m_IDPCA)
                            m_PCA.GrabarNombreIcono p_Error
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                            Set m_PCA = Nothing
                        Next
                    End If
                End With
                Set m_PC = Nothing
            Next
        End If
        Set m_Riesgo = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarSiHayErroresYNombreNodosGral ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function RellenarNombreIconoRiesgosActivos(Optional ByRef p_Error As String) As String

    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Riesgo As riesgo
    Dim m_ColPlanes As Scripting.Dictionary
    Dim m_IDPlan As Variant
    Dim m_Plan As Object
    
    
    On Error GoTo errores
    Set m_Col = Constructor.getRiesgosActivos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Riesgo = m_Col(m_ID)
        With m_Riesgo
            VBA.DoEvents
            Debug.Print .CodigoUnicoCalculado
            VBA.DoEvents
            'If .NombreIcono = "" Then
                .GrabarNombreIcono p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            'End If
           
            
            'If .CodigoUnicoCalculado = "099R002" Then Stop
            'If .HayErrorEnRiesgo = "" Then
                .HayErrorEnRiesgoGrabar p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            'End If
            'If .NombreNodoDesc = "" Then
                .GrabarNombreNodoDesc p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            'End If
            'If .NombreNodoEstado = "" Then
                .GrabarNombreNodoEstado p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            'End If
            'If .ColorIcono = "" Then
                .GrabarColorIcono p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            'End If
        End With
        
        Set m_ColPlanes = m_Riesgo.ColPMs
        p_Error = m_Riesgo.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ColPlanes Is Nothing Then
            For Each m_IDPlan In m_ColPlanes
                Set m_Plan = m_ColPlanes(m_IDPlan)
                With m_Plan
                    'If .NombreIcono = "" Then
                        .GrabarNombreIcono p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    'End If
                End With
                Set m_Plan = Nothing
            Next
        End If
        Set m_ColPlanes = m_Riesgo.ColPCs
        p_Error = m_Riesgo.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ColPlanes Is Nothing Then
            For Each m_IDPlan In m_ColPlanes
                Set m_Plan = m_ColPlanes(m_IDPlan)
                With m_Plan
                    'If .NombreIcono = "" Then
                        .GrabarNombreIcono p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    'End If
                End With
                Set m_Plan = Nothing
            Next
        End If
        Set m_Riesgo = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarNombreIconoRiesgosActivos ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function RellenarCampoEstadoEnActivos(Optional ByRef p_Error As String) As String

    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Riesgo As riesgo
    Dim m_ColPlanes As Scripting.Dictionary
    Dim m_IDPlan As Variant
    Dim m_Plan As Object
    Dim m_valorInicial As String
    Dim m_valorFinal As String
    
    On Error GoTo errores
    Set m_Col = Constructor.getRiesgosActivos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Riesgo = m_Col(m_ID)
        With m_Riesgo
            VBA.DoEvents
            
            VBA.DoEvents
            m_valorInicial = .Estado
            m_valorFinal = .ESTADOCalculadoTexto
            If m_valorInicial <> m_valorFinal Then
                .GrabarEstado p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                Debug.Print .CodigoUnicoCalculado, m_valorInicial & "|" & m_valorFinal
            Else
                Debug.Print .CodigoUnicoCalculado, "IGUAL|" & m_valorInicial
            End If
           
            
            
        End With
        
        Set m_ColPlanes = m_Riesgo.ColPMs
        p_Error = m_Riesgo.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ColPlanes Is Nothing Then
            For Each m_IDPlan In m_ColPlanes
                Set m_Plan = m_ColPlanes(m_IDPlan)
                With m_Plan
                    m_valorInicial = .Estado
                    m_valorFinal = .ESTADOCalculadoTexto
                    If m_valorInicial <> m_valorFinal Then
                        .GrabarEstado p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        Debug.Print vbTab & m_Riesgo.CodigoUnicoCalculado & "|PM", m_valorInicial & "|" & m_valorFinal
                    Else
                        Debug.Print vbTab & m_Riesgo.CodigoUnicoCalculado & "|PM", "IGUAL|" & m_valorInicial
                    End If
                End With
                Set m_Plan = Nothing
            Next
        End If
        Set m_ColPlanes = m_Riesgo.ColPCs
        p_Error = m_Riesgo.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ColPlanes Is Nothing Then
            For Each m_IDPlan In m_ColPlanes
                Set m_Plan = m_ColPlanes(m_IDPlan)
                With m_Plan
                     m_valorInicial = .Estado
                    m_valorFinal = .ESTADOCalculadoTexto
                    If m_valorInicial <> m_valorFinal Then
                        .GrabarEstado p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        Debug.Print vbTab & m_Riesgo.CodigoUnicoCalculado & "|PC", m_valorInicial & "|" & m_valorFinal
                    Else
                        Debug.Print vbTab & m_Riesgo.CodigoUnicoCalculado & "|PC", "IGUAL|" & m_valorInicial
                    End If
                End With
                Set m_Plan = Nothing
            Next
        End If
        Set m_Riesgo = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarCampoEstadoEnActivos ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function RellenarNombreProyecto(Optional ByRef p_Error As String) As String
    
    Dim m_ID As Variant
    Dim m_Proyecto As Proyecto
    Dim m_SQL As String
    Dim m_NombreProyectoCalculado As String
    Dim m_NombreAlInicio As String
    
    On Error GoTo errores
    For Each m_ID In m_ObjEntorno.ColProyectosTotales
       ' If CStr(m_ID) = "98" Then Stop
        Set m_Proyecto = m_ObjEntorno.ColProyectosTotales(m_ID)
        m_NombreAlInicio = m_Proyecto.NombreProyecto
        m_NombreProyectoCalculado = m_Proyecto.NombreProyectoCalculado
        If m_NombreProyectoCalculado = "" Then
            Debug.Print m_Proyecto.IDProyecto, "SIN NombreProyectoCalculado"
        End If
        m_Proyecto.NombreProyecto = m_NombreProyectoCalculado
        If m_Proyecto.NombreProyecto <> m_NombreAlInicio Then
            If m_Proyecto.NombreProyecto <> "" Then
                 m_SQL = "UPDATE TbProyectos SET NombreProyecto ='" & m_Proyecto.NombreProyecto & "' " & _
                        "WHERE IDProyecto=" & m_ID & ";"
            
            End If
           
            getdb().Execute m_SQL
            VBA.DoEvents
            Debug.Print m_Proyecto.IDProyecto, m_NombreAlInicio, m_Proyecto.NombreProyecto
            VBA.DoEvents
        Else
            VBA.DoEvents
            Debug.Print m_Proyecto.IDProyecto, "IGUAL"
            VBA.DoEvents
        End If
siguiente:
        Set m_Proyecto = Nothing
    Next
    
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarNombreProyecto ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function RellenarCorreoRacs(Optional ByRef p_Error As String) As String
    
    Dim m_ID As Variant
    Dim m_Proyecto As Proyecto
    Dim m_SQL As String
    Dim m_CorreoRacCalculado As String
    Dim m_CorreoRacAlInicio As String
    
    On Error GoTo errores
    For Each m_ID In m_ObjEntorno.ColProyectosTotales
       ' If CStr(m_ID) = "98" Then Stop
        Set m_Proyecto = m_ObjEntorno.ColProyectosTotales(m_ID)
        m_CorreoRacAlInicio = m_Proyecto.CorreoRAC
        m_CorreoRacCalculado = m_Proyecto.CorreoRACCalculado
        
        If m_CorreoRacCalculado = "" Then
            Debug.Print m_Proyecto.IDProyecto, "SIN CorreoRACCalculado"
            GoTo siguiente
        End If
        m_Proyecto.CorreoRAC = m_CorreoRacCalculado
        If m_Proyecto.CorreoRAC <> m_CorreoRacAlInicio Then
            If m_Proyecto.CorreoRAC <> "" Then
                m_SQL = "UPDATE TbProyectos SET CorreoRAC ='" & m_Proyecto.CorreoRAC & "' " & _
                        "WHERE IDProyecto=" & m_ID & ";"
            Else
                m_SQL = "UPDATE TbProyectos SET CorreoRAC =Null " & _
                        "WHERE IDProyecto=" & m_ID & ";"
            End If
            
           
            getdb().Execute m_SQL
            VBA.DoEvents
            Debug.Print m_Proyecto.IDProyecto, m_CorreoRacAlInicio, m_Proyecto.CorreoRAC
            VBA.DoEvents
        Else
            VBA.DoEvents
            Debug.Print m_Proyecto.IDProyecto, "IGUAL"
            VBA.DoEvents
        End If
siguiente:
        Set m_Proyecto = Nothing
    Next
    
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarCorreoRacs ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function RellenarFechaPrevistaCierre(Optional ByRef p_Error As String) As String
    
    Dim m_ID As Variant
    Dim m_Proyecto As Proyecto
    Dim m_SQL As String
    Dim m_FechaPrevistaCierreCalculado As String
    Dim m_FechaPrevistaCierreAlInicio As String
    
    On Error GoTo errores
    For Each m_ID In m_ObjEntorno.ColProyectosTotales
       ' If CStr(m_ID) = "98" Then Stop
        Set m_Proyecto = m_ObjEntorno.ColProyectosTotales(m_ID)
        m_FechaPrevistaCierreAlInicio = m_Proyecto.FechaPrevistaCierre
        m_FechaPrevistaCierreCalculado = m_Proyecto.FechaPrevistaCierreCalculada
        If m_FechaPrevistaCierreCalculado = "" Then
            Debug.Print m_Proyecto.IDProyecto, "SIN m_FechaPrevistaCierreCalculado"
            GoTo siguiente
        End If
        m_Proyecto.FechaPrevistaCierre = m_FechaPrevistaCierreCalculado
        If m_Proyecto.FechaPrevistaCierre <> m_FechaPrevistaCierreAlInicio Then
            If m_Proyecto.FechaPrevistaCierre <> "" Then
                 m_SQL = "UPDATE TbProyectos SET FechaPrevistaCierre =#" & Format(m_Proyecto.FechaPrevistaCierre, "mm/dd/yyyy") & "# " & _
                        "WHERE IDProyecto=" & m_ID & ";"
            
            End If
           
            getdb().Execute m_SQL
            VBA.DoEvents
            Debug.Print m_Proyecto.IDProyecto, m_FechaPrevistaCierreAlInicio, m_Proyecto.FechaPrevistaCierre
            VBA.DoEvents
        Else
            VBA.DoEvents
            Debug.Print m_Proyecto.IDProyecto, "IGUAL"
            VBA.DoEvents
        End If
siguiente:
        Set m_Proyecto = Nothing
    Next
    
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarFechaPrevistaCierre ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function RellenarFechaProxPublicacion(Optional ByRef p_Error As String) As String
    
    Dim m_ID As Variant
    Dim m_Edicion As Edicion
    Dim m_SQL As String
    Dim m_FechaMaxProximaPublicacionCalculado As String
    Dim m_FechaMaxProximaPublicacionAlInicio As String
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    Set m_Col = Constructor.getEdicionesActivas(p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
       ' If CStr(m_ID) = "98" Then Stop
        Set m_Edicion = m_Col(m_ID)
        m_FechaMaxProximaPublicacionAlInicio = m_Edicion.FechaMaxProximaPublicacion
        m_FechaMaxProximaPublicacionCalculado = m_Edicion.FechaMaxProximaPublicacionCalculado
        
        If m_FechaMaxProximaPublicacionCalculado = "" Then
            Debug.Print m_Edicion.IDProyecto, "SIN m_FechaMaxProximaPublicacionCalculado"
            GoTo siguiente
        End If
        m_Edicion.FechaMaxProximaPublicacion = m_FechaMaxProximaPublicacionCalculado
        If m_Edicion.FechaMaxProximaPublicacion <> m_FechaMaxProximaPublicacionAlInicio Then
            If m_Edicion.FechaMaxProximaPublicacion <> "" Then
                 m_SQL = "UPDATE TbProyectosEdiciones SET FechaMaxProximaPublicacion =#" & Format(m_Edicion.FechaMaxProximaPublicacion, "mm/dd/yyyy") & "# " & _
                        "WHERE IDEdicion=" & m_ID & ";"
            
            End If
           
            getdb().Execute m_SQL
            VBA.DoEvents
            Debug.Print m_Edicion.IDProyecto, m_FechaMaxProximaPublicacionAlInicio, m_Edicion.FechaMaxProximaPublicacion
            VBA.DoEvents
        Else
            VBA.DoEvents
            Debug.Print m_Edicion.IDProyecto, "IGUAL"
            VBA.DoEvents
        End If
siguiente:
        Set m_Edicion = Nothing
    Next
    
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarFechaProxPublicacion ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function


