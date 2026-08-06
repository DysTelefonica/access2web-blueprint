Attribute VB_Name = "Constructor"
Option Compare Database
Option Explicit

' Counter for query-uniqueness test in B2 (Punto 15)
Private m_ContadorLlamadasGetAnexosTotalesDeEdicion As Long
Private m_ContadorLlamadasGetAnexosTotalesDeRiesgo As Long

' ============================================================
' Batch-loading helpers for CargarArbol N+1 elimination
' ============================================================

Private Function JoinCollectionIDs(p_Col As Collection, p_Delim As String) As String
    Dim i As Long
    JoinCollectionIDs = ""
    For i = 1 To p_Col.Count
        If i > 1 Then
            JoinCollectionIDs = JoinCollectionIDs & p_Delim
        End If
        JoinCollectionIDs = JoinCollectionIDs & CStr(p_Col(i))
    Next i
End Function

Public Function getPMsPorEdicionBatch( _
                            p_IDEdicion As String, _
                            p_lista_IDRiesgos As Collection, _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary
    ' Batch-loads all PMs for a list of riesgos in a single DAO query.
    ' Returns Dictionary keyed by IDRiesgo -> inner Dictionary keyed by IDMitigacion -> PM object.
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_PM As pm
    Dim m_IdRiesgo As String
    Dim m_IDMitigacion As String
    Dim colPMsPorRiesgo As Scripting.Dictionary
    Dim m_DicPMsBatch As Scripting.Dictionary
    
    On Error GoTo errores
    
    If p_IDEdicion = "" Then
        p_Error = "Falta p_IDEdicion"
        Err.Raise 1000
    End If
    If p_lista_IDRiesgos Is Nothing Then
        Exit Function
    End If
    If p_lista_IDRiesgos.Count = 0 Then
        Exit Function
    End If
    
    Dim listaIDs As String
    listaIDs = JoinCollectionIDs(p_lista_IDRiesgos, ",")
    If listaIDs = "" Then
        Exit Function
    End If
    
    m_SQL = "SELECT TbRiesgosPlanMitigacionPpal.* " & _
            "FROM TbRiesgosPlanMitigacionPpal " & _
            "WHERE IDRiesgo In(" & listaIDs & ") " & _
            "ORDER BY IDRiesgo, IDMitigacion;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_IdRiesgo = Nz(.fields("IDRiesgo").value, "")
            m_IDMitigacion = Nz(.fields("IDMitigacion").value, "")
            If m_IDMitigacion <> "" Then
                Set m_PM = New pm
                For Each m_Campo In m_PM.ColCampos
                    m_PM.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If m_DicPMsBatch Is Nothing Then
                    Set m_DicPMsBatch = New Scripting.Dictionary
                    m_DicPMsBatch.CompareMode = TextCompare
                End If
                If Not m_DicPMsBatch.Exists(m_IdRiesgo) Then
                    Set colPMsPorRiesgo = New Scripting.Dictionary
                    colPMsPorRiesgo.CompareMode = TextCompare
                    m_DicPMsBatch.Add m_IdRiesgo, colPMsPorRiesgo
                Else
                    Set colPMsPorRiesgo = m_DicPMsBatch(m_IdRiesgo)
                End If
                If Not colPMsPorRiesgo.Exists(m_IDMitigacion) Then
                    colPMsPorRiesgo.Add m_IDMitigacion, m_PM
                End If
            End If
            Set m_PM = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set getPMsPorEdicionBatch = m_DicPMsBatch
    ' Also populate the global PM cache so GetCachedPM hits it instead of the DB
    If Not m_DicPMsBatch Is Nothing Then
        Dim keyIdRiesgo As Variant
        Dim colPMs As Scripting.Dictionary
        Dim keyIdPM As Variant
        Dim objPM As pm
        For Each keyIdRiesgo In m_DicPMsBatch
            Set colPMs = m_DicPMsBatch(keyIdRiesgo)
            For Each keyIdPM In colPMs
                Set objPM = colPMs(keyIdPM)
                Call GetCachedPM(keyIdPM) ' populate cache via side-effect
            Next
        Next
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getPMsPorEdicionBatch ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPCsPorEdicionBatch( _
                            p_IDEdicion As String, _
                            p_lista_IDRiesgos As Collection, _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary
    ' Batch-loads all PCs for a list of riesgos in a single DAO query.
    ' Returns Dictionary keyed by IDRiesgo -> inner Dictionary keyed by IDContingencia -> PC object.
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_PC As pc
    Dim m_IdRiesgo As String
    Dim m_IDContingencia As String
    Dim colPCsPorRiesgo As Scripting.Dictionary
    Dim m_DicPCsBatch As Scripting.Dictionary
    
    On Error GoTo errores
    
    If p_IDEdicion = "" Then
        p_Error = "Falta p_IDEdicion"
        Err.Raise 1000
    End If
    If p_lista_IDRiesgos Is Nothing Then
        Exit Function
    End If
    If p_lista_IDRiesgos.Count = 0 Then
        Exit Function
    End If
    
    Dim listaIDs As String
    listaIDs = JoinCollectionIDs(p_lista_IDRiesgos, ",")
    If listaIDs = "" Then
        Exit Function
    End If
    
    m_SQL = "SELECT TbRiesgosPlanContingenciaPpal.* " & _
            "FROM TbRiesgosPlanContingenciaPpal " & _
            "WHERE IDRiesgo In(" & listaIDs & ") " & _
            "ORDER BY IDRiesgo, IDContingencia;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_IdRiesgo = Nz(.fields("IDRiesgo").value, "")
            m_IDContingencia = Nz(.fields("IDContingencia").value, "")
            If m_IDContingencia <> "" Then
                Set m_PC = New pc
                For Each m_Campo In m_PC.ColCampos
                    m_PC.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If m_DicPCsBatch Is Nothing Then
                    Set m_DicPCsBatch = New Scripting.Dictionary
                    m_DicPCsBatch.CompareMode = TextCompare
                End If
                If Not m_DicPCsBatch.Exists(m_IdRiesgo) Then
                    Set colPCsPorRiesgo = New Scripting.Dictionary
                    colPCsPorRiesgo.CompareMode = TextCompare
                    m_DicPCsBatch.Add m_IdRiesgo, colPCsPorRiesgo
                Else
                    Set colPCsPorRiesgo = m_DicPCsBatch(m_IdRiesgo)
                End If
                If Not colPCsPorRiesgo.Exists(m_IDContingencia) Then
                    colPCsPorRiesgo.Add m_IDContingencia, m_PC
                End If
            End If
            Set m_PC = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set getPCsPorEdicionBatch = m_DicPCsBatch
    ' Also populate the global PC cache so GetCachedPC hits it instead of the DB
    If Not m_DicPCsBatch Is Nothing Then
        Dim keyIdRiesgoPC As Variant
        Dim colPCs As Scripting.Dictionary
        Dim keyIdPC As Variant
        Dim objPC As pc
        For Each keyIdRiesgoPC In m_DicPCsBatch
            Set colPCs = m_DicPCsBatch(keyIdRiesgoPC)
            For Each keyIdPC In colPCs
                Set objPC = colPCs(keyIdPC)
                Call GetCachedPC(keyIdPC) ' populate cache via side-effect
            Next
        Next
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getPCsPorEdicionBatch ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPMAccionesBatch( _
                            p_lista_IDMitigacion As Collection, _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary
    ' Batch-loads all PM actions for a list of mitigations in a single DAO query.
    ' Returns Dictionary keyed by IDMitigacion -> inner Dictionary keyed by IDAccionMitigacion -> PMAccion object.
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_PMAccion As PMAccion
    Dim m_IDMitigacion As String
    Dim m_IDAccion As String
    Dim colAccionesPorMitigacion As Scripting.Dictionary
    Dim m_DicPMAccionesBatch As Scripting.Dictionary
    
    On Error GoTo errores
    
    If p_lista_IDMitigacion Is Nothing Then
        Exit Function
    End If
    If p_lista_IDMitigacion.Count = 0 Then
        Exit Function
    End If
    
    Dim listaIDs As String
    listaIDs = JoinCollectionIDs(p_lista_IDMitigacion, ",")
    If listaIDs = "" Then
        Exit Function
    End If
    
    m_SQL = "SELECT TbRiesgosPlanMitigacionDetalle.* " & _
            "FROM TbRiesgosPlanMitigacionDetalle " & _
            "WHERE IDMitigacion In(" & listaIDs & ") " & _
            "ORDER BY IDMitigacion, IDAccionMitigacion;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_IDMitigacion = Nz(.fields("IDMitigacion").value, "")
            m_IDAccion = Nz(.fields("IDAccionMitigacion").value, "")
            If m_IDAccion <> "" Then
                Set m_PMAccion = New PMAccion
                For Each m_Campo In m_PMAccion.ColCampos
                    m_PMAccion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If m_DicPMAccionesBatch Is Nothing Then
                    Set m_DicPMAccionesBatch = New Scripting.Dictionary
                    m_DicPMAccionesBatch.CompareMode = TextCompare
                End If
                If Not m_DicPMAccionesBatch.Exists(m_IDMitigacion) Then
                    Set colAccionesPorMitigacion = New Scripting.Dictionary
                    colAccionesPorMitigacion.CompareMode = TextCompare
                    m_DicPMAccionesBatch.Add m_IDMitigacion, colAccionesPorMitigacion
                Else
                    Set colAccionesPorMitigacion = m_DicPMAccionesBatch(m_IDMitigacion)
                End If
                If Not colAccionesPorMitigacion.Exists(m_IDAccion) Then
                    colAccionesPorMitigacion.Add m_IDAccion, m_PMAccion
                End If
            End If
            Set m_PMAccion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set getPMAccionesBatch = m_DicPMAccionesBatch
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getPMAccionesBatch ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPCAccionesBatch( _
                            p_lista_IDContingencia As Collection, _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary
    ' Batch-loads all PC actions for a list of contingencias in a single DAO query.
    ' Returns Dictionary keyed by IDContingencia -> inner Dictionary keyed by IDAccionContingencia -> PCAccion object.
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_PCAccion As PCAccion
    Dim m_IDContingencia As String
    Dim m_IDAccion As String
    Dim colAccionesPorContingencia As Scripting.Dictionary
    Dim m_DicPCAccionesBatch As Scripting.Dictionary
    
    On Error GoTo errores
    
    If p_lista_IDContingencia Is Nothing Then
        Exit Function
    End If
    If p_lista_IDContingencia.Count = 0 Then
        Exit Function
    End If
    
    Dim listaIDs As String
    listaIDs = JoinCollectionIDs(p_lista_IDContingencia, ",")
    If listaIDs = "" Then
        Exit Function
    End If
    
    m_SQL = "SELECT TbRiesgosPlanContingenciaDetalle.* " & _
            "FROM TbRiesgosPlanContingenciaDetalle " & _
            "WHERE IDContingencia In(" & listaIDs & ") " & _
            "ORDER BY IDContingencia, IDAccionContingencia;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_IDContingencia = Nz(.fields("IDContingencia").value, "")
            m_IDAccion = Nz(.fields("IDAccionContingencia").value, "")
            If m_IDAccion <> "" Then
                Set m_PCAccion = New PCAccion
                For Each m_Campo In m_PCAccion.ColCampos
                    m_PCAccion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If m_DicPCAccionesBatch Is Nothing Then
                    Set m_DicPCAccionesBatch = New Scripting.Dictionary
                    m_DicPCAccionesBatch.CompareMode = TextCompare
                End If
                If Not m_DicPCAccionesBatch.Exists(m_IDContingencia) Then
                    Set colAccionesPorContingencia = New Scripting.Dictionary
                    colAccionesPorContingencia.CompareMode = TextCompare
                    m_DicPCAccionesBatch.Add m_IDContingencia, colAccionesPorContingencia
                Else
                    Set colAccionesPorContingencia = m_DicPCAccionesBatch(m_IDContingencia)
                End If
                If Not colAccionesPorContingencia.Exists(m_IDAccion) Then
                    colAccionesPorContingencia.Add m_IDAccion, m_PCAccion
                End If
            End If
            Set m_PCAccion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set getPCAccionesBatch = m_DicPCAccionesBatch
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getPCAccionesBatch ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getUsuarioConectadoPorMaquina( _
                                                Optional ByRef p_Error As String _
                                                ) As Usuario
    Dim objNetwork As Object
    On Error GoTo errores
    Set objNetwork = CreateObject("Wscript.Network")
    Set getUsuarioConectadoPorMaquina = Constructor.getUsuario(, objNetwork.UserName, , , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set objNetwork = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioConectadoPorMaquina ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getEdicion( _
                            p_IDEdicion As String, _
                            Optional ByRef p_Error As String _
                            ) As Edicion


    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Result As Edicion  ' obs #16137 fix: local var, asignada al final

    On Error GoTo errores

    If p_IDEdicion = "" Then
        p_Error = "Fatalta el p_IDEdicion"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * " & _
            "FROM TbProyectosEdiciones " & _
            "WHERE IDEdicion=" & p_IDEdicion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set m_Result = New Edicion
        For Each m_Campo In m_Result.ColCampos
            m_Result.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set getEdicion = m_Result
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getEdicion ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getEdicionV2( _
                            p_IDProyecto As String, _
                            p_Edicion As String, _
                            Optional ByRef p_Error As String _
                            ) As Edicion
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDProyecto = "" Then
        Exit Function
    End If
    If p_Edicion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbProyectosEdiciones " & _
            "WHERE IDProyecto=" & p_IDProyecto & "AND Edicion=" & p_Edicion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getEdicionV2 = New Edicion
        For Each m_Campo In getEdicionV2.ColCampos
            getEdicionV2.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getEdicionV2 ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getLogPublicacion( _
                                    p_Id As String, _
                                    Optional ByRef p_Error As String _
                                    ) As PublicacionLog
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_Id = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbLogPublicaciones " & _
                "WHERE ID=" & p_Id & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getLogPublicacion = New PublicacionLog
        For Each m_Campo In getLogPublicacion.ColCampos
            getLogPublicacion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getLogPublicacion ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getLogPublicaciones( _
                                        p_IDEdicion As String, _
                                        Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_PublicacionLog As PublicacionLog
    
    On Error GoTo errores
    
    If p_IDEdicion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbLogPublicaciones " & _
                "WHERE IDEdicion=" & p_IDEdicion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_PublicacionLog = New PublicacionLog
            For Each m_Campo In m_PublicacionLog.ColCampos
                m_PublicacionLog.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getLogPublicaciones Is Nothing Then
                Set getLogPublicaciones = New Scripting.Dictionary
                getLogPublicaciones.CompareMode = TextCompare
            End If
            If Not getLogPublicaciones.Exists(m_PublicacionLog.ID) Then
                getLogPublicaciones.Add m_PublicacionLog.ID, m_PublicacionLog
            End If
            Set m_PublicacionLog = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getLogPublicaciones ha devuelto el error: " & Err.description
    End If
End Function

Public Function getSQLLogPublicacionesProyecto( _
                                                p_IDProyecto As String, _
                                                Optional ByRef p_Error As String _
                                            ) As String

    On Error GoTo errores

    p_Error = ""
    If p_IDProyecto = "" Then
        Exit Function
    End If

    getSQLLogPublicacionesProyecto = "SELECT TbLogPublicaciones.* FROM TbLogPublicaciones " & _
        "INNER JOIN TbProyectosEdiciones ON TbLogPublicaciones.IDEdicion = TbProyectosEdiciones.IDEdicion " & _
        "WHERE TbProyectosEdiciones.IDProyecto=" & p_IDProyecto & " " & _
        "ORDER BY TbLogPublicaciones.FechaRegistro, TbLogPublicaciones.ID;"
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getSQLLogPublicacionesProyecto ha devuelto el error: " & Err.description
    End If
End Function

Public Function getLogPublicacionesProyecto( _
                                                p_IDProyecto As String, _
                                                Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_PublicacionLog As PublicacionLog

    On Error GoTo errores

    If p_IDProyecto = "" Then
        Exit Function
    End If
    m_SQL = getSQLLogPublicacionesProyecto(p_IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_SQL = "" Then
        Exit Function
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_PublicacionLog = New PublicacionLog
            For Each m_Campo In m_PublicacionLog.ColCampos
                m_PublicacionLog.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getLogPublicacionesProyecto Is Nothing Then
                Set getLogPublicacionesProyecto = New Scripting.Dictionary
                getLogPublicacionesProyecto.CompareMode = TextCompare
            End If
            If Not getLogPublicacionesProyecto.Exists(m_PublicacionLog.ID) Then
                getLogPublicacionesProyecto.Add m_PublicacionLog.ID, m_PublicacionLog
            End If
            Set m_PublicacionLog = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getLogPublicacionesProyecto ha devuelto el error: " & Err.description
    End If
End Function
Public Function getLogCorreosResponsables( _
                                        p_IDEdicion As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_EdicionCorreoRevision As EdicionCorreoRevision
    
    On Error GoTo errores
    
    If p_IDEdicion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbProyectoEdicionesCorreoRevision " & _
                "WHERE IDEdicion=" & p_IDEdicion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_EdicionCorreoRevision = New EdicionCorreoRevision
            For Each m_Campo In m_EdicionCorreoRevision.ColCampos
                m_EdicionCorreoRevision.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getLogCorreosResponsables Is Nothing Then
                Set getLogCorreosResponsables = New Scripting.Dictionary
                getLogCorreosResponsables.CompareMode = TextCompare
            End If
            If Not getLogCorreosResponsables.Exists(m_EdicionCorreoRevision.IDEnvioCorreoTecnico) Then
                getLogCorreosResponsables.Add m_EdicionCorreoRevision.IDEnvioCorreoTecnico, m_EdicionCorreoRevision
            End If
            Set m_EdicionCorreoRevision = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getLogCorreosResponsables ha devuelto el error: " & Err.description
    End If
End Function
Public Function getSQLLogCorreosResponsablesProyecto( _
                                                    p_IDProyecto As String, _
                                                    Optional ByRef p_Error As String _
                                                ) As String

    On Error GoTo errores

    p_Error = ""
    If p_IDProyecto = "" Then
        Exit Function
    End If

    getSQLLogCorreosResponsablesProyecto = "SELECT TbProyectoEdicionesCorreoRevision.* FROM TbProyectoEdicionesCorreoRevision " & _
        "INNER JOIN TbProyectosEdiciones ON TbProyectoEdicionesCorreoRevision.IDEdicion = TbProyectosEdiciones.IDEdicion " & _
        "WHERE TbProyectosEdiciones.IDProyecto=" & p_IDProyecto & " " & _
        "ORDER BY TbProyectoEdicionesCorreoRevision.FechaCorreoRevision, TbProyectoEdicionesCorreoRevision.IDEdicion, TbProyectoEdicionesCorreoRevision.IDEnvioCorreoTecnico;"
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getSQLLogCorreosResponsablesProyecto ha devuelto el error: " & Err.description
    End If
End Function

Public Function getLogCorreosResponsablesProyecto( _
                                                p_IDProyecto As String, _
                                                Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_EdicionCorreoRevision As EdicionCorreoRevision

    On Error GoTo errores

    If p_IDProyecto = "" Then
        Exit Function
    End If
    m_SQL = getSQLLogCorreosResponsablesProyecto(p_IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_SQL = "" Then
        Exit Function
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_EdicionCorreoRevision = New EdicionCorreoRevision
            For Each m_Campo In m_EdicionCorreoRevision.ColCampos
                m_EdicionCorreoRevision.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getLogCorreosResponsablesProyecto Is Nothing Then
                Set getLogCorreosResponsablesProyecto = New Scripting.Dictionary
                getLogCorreosResponsablesProyecto.CompareMode = TextCompare
            End If
            If Not getLogCorreosResponsablesProyecto.Exists(m_EdicionCorreoRevision.IDEnvioCorreoTecnico) Then
                getLogCorreosResponsablesProyecto.Add m_EdicionCorreoRevision.IDEnvioCorreoTecnico, m_EdicionCorreoRevision
            End If
            Set m_EdicionCorreoRevision = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getLogCorreosResponsablesProyecto ha devuelto el error: " & Err.description
    End If
End Function
Public Function getEdicionActiva( _
                                    p_IDProyecto As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Edicion
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_SQLLimitante As String
    
    On Error GoTo errores
    
    If p_IDProyecto = "" Then
        p_Error = "Fatalta el IDProyecto"
        Err.Raise 1000
    End If
    m_SQLLimitante = "SELECT TbProyectosEdiciones.IDEdicion " & _
                        "FROM TbProyectosEdiciones " & _
                        "WHERE (((TbProyectosEdiciones.IDProyecto)=" & p_IDProyecto & _
                        ") AND ((TbProyectosEdiciones.FechaPublicacion) Is Null));"
    
    m_SQL = "SELECT * FROM TbProyectosEdiciones " & _
                "WHERE IDEdicion In(" & m_SQLLimitante & ");"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getEdicionActiva = New Edicion
        For Each m_Campo In getEdicionActiva.ColCampos
            getEdicionActiva.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getEdicionActiva ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function


Public Function getEdicionUltima( _
                                p_IDProyecto As String, _
                                Optional ByRef p_Error As String _
                                ) As Edicion
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_SQLLimitante As String
    
    
    On Error GoTo errores
    p_Error = ""
    If p_IDProyecto = "" Then
        p_Error = "Fatalta el IDProyecto"
        Err.Raise 1000
    End If
    m_SQLLimitante = "SELECT Max(IDEdicion) AS MáxDeIDEdicion " & _
                    "FROM TbProyectosEdiciones " & _
                    "WHERE IDProyecto=" & p_IDProyecto & ""
    m_SQL = "SELECT * FROM TbProyectosEdiciones " & _
                "WHERE IDEdicion In(" & m_SQLLimitante & ");"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getEdicionUltima = New Edicion
        For Each m_Campo In getEdicionUltima.ColCampos
            getEdicionUltima.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getEdicionUltima ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getEdicionUltimaPublicada( _
                                        p_IDProyecto As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Edicion
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_SQLLimitante As String
    
    
    On Error GoTo errores
    p_Error = ""
    If p_IDProyecto = "" Then
        p_Error = "Fatalta el IDProyecto"
        Err.Raise 1000
    End If
    m_SQLLimitante = "SELECT Max(IDEdicion) AS MáxDeIDEdicion " & _
                    "FROM TbProyectosEdiciones " & _
                    "WHERE IDProyecto=" & p_IDProyecto & _
                    " AND Not FechaPublicacion Is Null;"
    m_SQL = "SELECT * FROM TbProyectosEdiciones " & _
                "WHERE IDEdicion In(" & m_SQLLimitante & ");"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getEdicionUltimaPublicada = New Edicion
        For Each m_Campo In getEdicionUltimaPublicada.ColCampos
            getEdicionUltimaPublicada.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getEdicionUltimaPublicada ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getEdicionPrimera( _
                                p_IDProyecto As String, _
                                Optional ByRef p_Error As String _
                                ) As Edicion
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDProyecto = "" Then
        p_Error = "Fatalta el p_IDProyecto"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * FROM TbProyectosEdiciones " & _
                "WHERE IDProyecto=" & p_IDProyecto & " ORDER BY IDEdicion;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getEdicionPrimera = New Edicion
        For Each m_Campo In getEdicionPrimera.ColCampos
            getEdicionPrimera.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getEdicionPrimera ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getEdicionSiguiente( _
                                    p_ObjEdicion As Edicion, _
                                    Optional ByRef p_Error As String _
                                    ) As Edicion
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
     
    If p_ObjEdicion Is Nothing Then
        p_Error = "Falta la edición"
        Err.Raise 1000
    End If
    If p_ObjEdicion.Edicion = "" Then
        p_Error = "Falta la edición"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * FROM TbProyectosEdiciones " & _
                "WHERE IDProyecto=" & p_ObjEdicion.IDProyecto & _
                "AND Edicion= " & CInt(p_ObjEdicion.Edicion) + 1 & " ;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getEdicionSiguiente = New Edicion
        For Each m_Campo In getEdicionSiguiente.ColCampos
            getEdicionSiguiente.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getEdicionSiguiente ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getEdicionAnterior( _
                                    p_ObjEdicion As Edicion, _
                                    Optional ByRef p_Error As String _
                                    ) As Edicion
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_ObjEdicion Is Nothing Then
        p_Error = "Falta la edición"
        Err.Raise 1000
    End If
    If p_ObjEdicion.Edicion = "" Then
        p_Error = "Falta la edición"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * " & _
            "FROM TbProyectosEdiciones " & _
            "WHERE IDProyecto=" & p_ObjEdicion.IDProyecto & _
            "AND Edicion= " & CInt(p_ObjEdicion.Edicion) - 1 & " ;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getEdicionAnterior = New Edicion
        For Each m_Campo In getEdicionAnterior.ColCampos
            getEdicionAnterior.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getEdicionAnterior ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getEdiciones( _
                                p_IDProyecto As String, _
                                Optional p_Creciente As EnumSiNo = EnumSiNo.Sí, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_objEdicion As Edicion
    
    On Error GoTo errores
    
    If p_IDProyecto = "" Then
        p_Error = "Falta la p_IDProyecto"
        Err.Raise 1000
    End If
    If p_Creciente = EnumSiNo.Sí Then
        m_SQL = "SELECT * " & _
                "FROM TbProyectosEdiciones " & _
                "WHERE IDProyecto=" & p_IDProyecto & " ORDER BY IDEdicion;"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbProyectosEdiciones " & _
                "WHERE IDProyecto=" & p_IDProyecto & " ORDER BY IDEdicion DESC;"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_objEdicion = New Edicion
            For Each m_Campo In m_objEdicion.ColCampos
                m_objEdicion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEdiciones Is Nothing Then
                Set getEdiciones = New Scripting.Dictionary
                getEdiciones.CompareMode = TextCompare
            End If
            If Not getEdiciones.Exists(m_objEdicion.IDEdicion) Then
                getEdiciones.Add m_objEdicion.IDEdicion, m_objEdicion
            End If
            Set m_objEdicion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getEdiciones ha devuelto el error: " & Err.description
    End If
End Function
Public Function getEdicionesInvolucradasParaInforme( _
                                                    p_IDEdicionInforme As String, _
                                                    Optional p_ObjProyecto As Proyecto, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_objEdicion As Edicion
    Dim m_idProyecto As String
    
    On Error GoTo errores
    
    If p_IDEdicionInforme = "" Then
        p_Error = "Falta la p_IDProyecto"
        Err.Raise 1000
    End If
    If p_ObjProyecto Is Nothing Then
        Set m_objEdicion = Constructor.getEdicion(p_IDEdicionInforme, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set p_ObjProyecto = m_objEdicion.Proyecto
        p_Error = m_objEdicion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If p_ObjProyecto Is Nothing Then
        p_Error = "No se ha podido obtener el proyecto"
        Err.Raise 1000
    End If
    m_idProyecto = p_ObjProyecto.IDProyecto
    Set m_objEdicion = Nothing
    m_SQL = "SELECT TbProyectosEdiciones.* " & _
            "FROM TbProyectosEdiciones " & _
            "WHERE (((TbProyectosEdiciones.IDProyecto)=" & m_idProyecto & _
            ") AND ((TbProyectosEdiciones.IDEdicion)<=" & p_IDEdicionInforme & "))ORDER BY IDEdicion;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_objEdicion = New Edicion
            For Each m_Campo In m_objEdicion.ColCampos
                m_objEdicion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEdicionesInvolucradasParaInforme Is Nothing Then
                Set getEdicionesInvolucradasParaInforme = New Scripting.Dictionary
                getEdicionesInvolucradasParaInforme.CompareMode = TextCompare
            End If
            If Not getEdicionesInvolucradasParaInforme.Exists(m_objEdicion.IDEdicion) Then
                getEdicionesInvolucradasParaInforme.Add m_objEdicion.IDEdicion, m_objEdicion
            End If
            Set m_objEdicion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getEdicionesInvolucradasParaInforme ha devuelto el error: " & Err.description
    End If
End Function
Public Function getProyecto( _
                            p_IDProyecto As String, _
                            Optional ByRef p_Error As String _
                            ) As Proyecto
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDProyecto = "" Then
        p_Error = "Falta la p_IDProyecto"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * " & _
            "FROM TbProyectos " & _
            "WHERE IDProyecto=" & p_IDProyecto & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getProyecto = New Proyecto
        For Each m_Campo In getProyecto.ColCampos
            getProyecto.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getProyecto ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getProyectoPorExpediente( _
                                        p_IDExpediente As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Proyecto
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDExpediente = "" Then
        p_Error = "Falta la p_IDExpediente"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * " & _
            "FROM TbProyectos " & _
            "WHERE IDExpediente=" & p_IDExpediente & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getProyectoPorExpediente = New Proyecto
        For Each m_Campo In getProyectoPorExpediente.ColCampos
            getProyectoPorExpediente.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getProyectoPorExpediente ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getProyectoPorCodigo( _
                                    p_Proyecto As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Proyecto
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_Proyecto = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbProyectos " & _
            "WHERE Proyecto='" & p_Proyecto & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getProyectoPorCodigo = New Proyecto
        For Each m_Campo In getProyectoPorCodigo.ColCampos
            getProyectoPorCodigo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getProyectoPorCodigo ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getRiesgosExternos( _
                                    p_IDProyecto As String, _
                                    Optional p_EnumOrigenRiesgoExterno As EnumOrigenRiesgoExterno, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgoExt As RiesgoExterno
    Dim m_Origen As String
    Dim m_Where As String
    
    On Error GoTo errores
    
    If p_IDProyecto = "" Then
        p_Error = "Falta la p_IDProyecto"
        Err.Raise 1000
    End If
    If p_EnumOrigenRiesgoExterno = EnumOrigenRiesgoExterno.Oferta Then
        m_Origen = "Oferta"
    ElseIf p_EnumOrigenRiesgoExterno = EnumOrigenRiesgoExterno.Subcontratista Then
        m_Origen = "Suministrador"
    ElseIf p_EnumOrigenRiesgoExterno = EnumOrigenRiesgoExterno.Pedido Then
        m_Origen = "Pedido"
    
    End If
    If m_Origen <> "" Then
        m_Where = "WHERE IDProyecto=" & p_IDProyecto & " AND Origen='" & m_Origen & "';"
    Else
        m_Where = "WHERE IDProyecto=" & p_IDProyecto & ";"
    End If
    m_SQL = "SELECT TbRiesgosAIntegrar.* " & _
            "FROM TbProyectosEdiciones INNER JOIN TbRiesgosAIntegrar ON " & _
            "TbProyectosEdiciones.IDEdicion = TbRiesgosAIntegrar.IDEdicion " & _
            m_Where
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgoExt = New RiesgoExterno
            For Each m_Campo In m_ObjRiesgoExt.ColCampos
                m_ObjRiesgoExt.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosExternos Is Nothing Then
                Set getRiesgosExternos = New Scripting.Dictionary
                getRiesgosExternos.CompareMode = TextCompare
            End If
            If Not getRiesgosExternos.Exists(m_ObjRiesgoExt.IDRiesgoExt) Then
                getRiesgosExternos.Add m_ObjRiesgoExt.IDRiesgoExt, m_ObjRiesgoExt
            End If
            Set m_ObjRiesgoExt = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosExternos ha devuelto el error: " & Err.description
    End If
End Function


Public Function getAnexosByProyecto( _
                                    p_IDProyecto As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
    
    On Error GoTo errores
    
    If p_IDProyecto = "" Then
        p_Error = "Falta la p_IDProyecto"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * " & _
            "FROM TbAnexos " & _
            "WHERE IDProyecto=" & p_IDProyecto & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosByProyecto Is Nothing Then
                Set getAnexosByProyecto = New Scripting.Dictionary
                getAnexosByProyecto.CompareMode = TextCompare
            End If
            If Not getAnexosByProyecto.Exists(m_ObjAnexo.IDAnexo) Then
                getAnexosByProyecto.Add m_ObjAnexo.IDAnexo, m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getAnexosByProyecto ha devuelto el error: " & Err.description
    End If
End Function

Public Function getAnexosDeEdicion( _
                                    p_IDEdicion As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
    
    On Error GoTo errores
    
    If p_IDEdicion = "" Then
        p_Error = "Falta la p_IDEdicion"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * " & _
            "FROM TbAnexos " & _
            "WHERE IDEdicion=" & p_IDEdicion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosDeEdicion Is Nothing Then
                Set getAnexosDeEdicion = New Scripting.Dictionary
                getAnexosDeEdicion.CompareMode = TextCompare
            End If
            If Not getAnexosDeEdicion.Exists(m_ObjAnexo.IDAnexo) Then
                getAnexosDeEdicion.Add m_ObjAnexo.IDAnexo, m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getAnexosDeEdicion ha devuelto el error: " & Err.description
    End If
End Function

Public Function getAnexosDeRiesgo( _
                                    p_IDRiesgo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
    
    On Error GoTo errores
    
    If p_IDRiesgo = "" Then
        p_Error = "Falta la p_IDRiesgo"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * " & _
            "FROM TbAnexos " & _
            "WHERE IDRiesgo=" & p_IDRiesgo & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosDeRiesgo Is Nothing Then
                Set getAnexosDeRiesgo = New Scripting.Dictionary
                getAnexosDeRiesgo.CompareMode = TextCompare
            End If
            If Not getAnexosDeRiesgo.Exists(m_ObjAnexo.IDAnexo) Then
                getAnexosDeRiesgo.Add m_ObjAnexo.IDAnexo, m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getAnexosDeRiesgo ha devuelto el error: " & Err.description
    End If
End Function

' ============================================================
' getAnexosTotalesDeEdicion — B2 / Punto 15
'   Query UNICA que trae anexos directos de la edicion + anexos
'   de los riesgos hijos. Sin N+1 (no iterar getAnexosDeRiesgo).
'   Orden: directos primero (Tipo=E), luego por IDRiesgo ASC.
' ============================================================
Public Function getAnexosTotalesDeEdicion( _
                                            p_IDEdicion As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo
    Dim m_Result As Scripting.Dictionary  ' obs #16137 fix: Dictionary local,
                                            ' nunca reasignar el nombre de funcion dentro del
                                            ' bucle (patron que provocaba "matriz es fija"
                                            ' porque mezclaba return value con variable
                                            ' local de acumulacion).

    On Error GoTo errores

    If Len(Trim$(p_IDEdicion)) = 0 Then
        p_Error = "getAnexosTotalesDeEdicion: p_IDEdicion esta vacio"
        Err.Raise 1000
    End If

    m_ContadorLlamadasGetAnexosTotalesDeEdicion = m_ContadorLlamadasGetAnexosTotalesDeEdicion + 1

    m_SQL = "SELECT IDAnexo, IDProyecto, IDEdicion, IDRiesgo, Titulo, NombreArchivo, " & _
            "FechaAnexo, EvidenciaUTE, EvidenciaSuministrador " & _
            "FROM TbAnexos " & _
            "WHERE IDEdicion=" & p_IDEdicion & " " & _
            "OR IDRiesgo IN (SELECT IDRiesgo FROM TbRiesgos WHERE IDEdicion=" & p_IDEdicion & ") " & _
            "ORDER BY IIF(IDEdicion=" & p_IDEdicion & ", 0, 1), " & _
            "IDRiesgo ASC, IDAnexo ASC;"

    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If m_Result Is Nothing Then
                Set m_Result = New Scripting.Dictionary
                m_Result.CompareMode = TextCompare
            End If
            If Not m_Result.Exists(m_ObjAnexo.IDAnexo) Then
                m_Result.Add m_ObjAnexo.IDAnexo, m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing

    ' Asignar el return value UNA vez al final (no dentro del bucle).
    Set getAnexosTotalesDeEdicion = m_Result
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo constructor.getAnexosTotalesDeEdicion ha devuelto el error: " & Err.Description
    End If
End Function

' --- Test-only accessors (NO usar en produccion) ---
' Contrato: contar las llamadas a getAnexosTotalesDeEdicion para
' que los tests puedan verificar que NO hay N+1.
Public Function GetContadorLlamadasGetAnexosTotalesDeEdicion() As Long
    GetContadorLlamadasGetAnexosTotalesDeEdicion = m_ContadorLlamadasGetAnexosTotalesDeEdicion
End Function

Public Sub ResetContadorLlamadasGetAnexosTotalesDeEdicion()
    m_ContadorLlamadasGetAnexosTotalesDeEdicion = 0
End Sub

' ============================================================
' getAnexosTotalesDeRiesgo — B2 / Punto 15 (espejo Riesgo)
'   Query UNICA que trae los anexos directos del riesgo.
'   WHERE IDRiesgo = X (sin N+1; el cache-first vive en
'   modFormRiesgoDocumentosHelper.getAnexosDeRiesgoCached).
'   Orden: IDAnexo ASC para determinismo en tests.
' ============================================================
Public Function getAnexosTotalesDeRiesgo( _
                                            p_IDRiesgo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAnexo As Anexo

    On Error GoTo errores

    If Len(Trim$(p_IDRiesgo)) = 0 Then
        p_Error = "getAnexosTotalesDeRiesgo: p_IDRiesgo esta vacio"
        Err.Raise 1000
    End If

    m_ContadorLlamadasGetAnexosTotalesDeRiesgo = m_ContadorLlamadasGetAnexosTotalesDeRiesgo + 1

    m_SQL = "SELECT IDAnexo, IDProyecto, IDEdicion, IDRiesgo, Titulo, NombreArchivo, " & _
            "FechaAnexo, EvidenciaUTE, EvidenciaSuministrador " & _
            "FROM TbAnexos " & _
            "WHERE IDRiesgo=" & p_IDRiesgo & " " & _
            "ORDER BY IDAnexo ASC;"

    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAnexo = New Anexo
            For Each m_Campo In m_ObjAnexo.ColCampos
                m_ObjAnexo.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAnexosTotalesDeRiesgo Is Nothing Then
                Set getAnexosTotalesDeRiesgo = New Scripting.Dictionary
                getAnexosTotalesDeRiesgo.CompareMode = TextCompare
            End If
            If Not getAnexosTotalesDeRiesgo.Exists(m_ObjAnexo.IDAnexo) Then
                getAnexosTotalesDeRiesgo.Add m_ObjAnexo.IDAnexo, m_ObjAnexo
            End If
            Set m_ObjAnexo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo constructor.getAnexosTotalesDeRiesgo ha devuelto el error: " & Err.Description
    End If
End Function

' --- Test-only accessors (NO usar en produccion) ---
' Contrato: contar las llamadas a getAnexosTotalesDeRiesgo para
' que los tests puedan verificar que NO hay N+1.
Public Function GetContadorLlamadasGetAnexosTotalesDeRiesgo() As Long
    GetContadorLlamadasGetAnexosTotalesDeRiesgo = m_ContadorLlamadasGetAnexosTotalesDeRiesgo
End Function

Public Sub ResetContadorLlamadasGetAnexosTotalesDeRiesgo()
    m_ContadorLlamadasGetAnexosTotalesDeRiesgo = 0
End Sub



Public Function getNombreArchivoAyuda( _
                                        p_NombreFormulario As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    If p_NombreFormulario = "" Then
        p_Error = "Se ha de indicar p_NombreFormulario"
        Exit Function
    End If
    m_SQL = "SELECT TbHerramientaDocAyuda.* " & _
            "FROM TbHerramientaDocAyuda " & _
            "WHERE NombreFormulario='" & p_NombreFormulario & "' ;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            getNombreArchivoAyuda = Nz(.fields("NombreArchivoAyuda"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getNombreArchivoAyuda ha devuelto el error: " & Err.description
    End If
End Function



Public Function getListaImpactosPorTipo( _
                                        p_Tipo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjAreaImpacto As AreaImpacto
    
    On Error GoTo errores
    If p_Tipo = "" Then
        p_Error = "Falta p_Tipo"
        Err.Raise 1000
    End If
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosAreasImpacto " & _
            "WHERE Tipo='" & p_Tipo & "' ORDER BY ordinal;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjAreaImpacto = New AreaImpacto
            For Each m_Campo In m_ObjAreaImpacto.ColCampos
                m_ObjAreaImpacto.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getListaImpactosPorTipo Is Nothing Then
                Set getListaImpactosPorTipo = New Scripting.Dictionary
                getListaImpactosPorTipo.CompareMode = TextCompare
            End If
            If Not getListaImpactosPorTipo.Exists(m_ObjAreaImpacto.Ordinal) Then
                getListaImpactosPorTipo.Add m_ObjAreaImpacto.Ordinal, m_ObjAreaImpacto
            End If
            Set m_ObjAreaImpacto = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getListaImpactosPorTipo ha devuelto el error: " & Err.description
    End If
End Function


Public Function getProyectos( _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Proyecto As Proyecto
    
    
    On Error GoTo errores
    
    
    m_SQL = "TbProyectos"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Proyecto = New Proyecto
            For Each m_Campo In m_Proyecto.ColCampos
                m_Proyecto.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getProyectos Is Nothing Then
                Set getProyectos = New Scripting.Dictionary
                getProyectos.CompareMode = TextCompare
            End If
            If Not getProyectos.Exists(m_Proyecto.IDProyecto) Then
                getProyectos.Add m_Proyecto.IDProyecto, m_Proyecto
            End If
            Set m_Proyecto = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getProyectos ha devuelto el error: " & Err.description
    End If
End Function
Public Function getProyectosDeUsuarios( _
                                        p_NombreUsuario As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim m_ID As Variant
    Dim m_Proyecto As Proyecto
    
    
    On Error GoTo errores
    If m_ObjEntorno.ColProyectosTotales Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ObjEntorno.ColProyectosTotales
        Set m_Proyecto = m_ObjEntorno.ColProyectosTotales(m_ID)
        If InStr(1, m_Proyecto, p_NombreUsuario) = 0 Then
            GoTo siguiente
        End If
        If getProyectosDeUsuarios Is Nothing Then
            Set getProyectosDeUsuarios = New Scripting.Dictionary
            getProyectosDeUsuarios.CompareMode = TextCompare
        End If
        If Not getProyectosDeUsuarios.Exists(m_Proyecto.IDProyecto) Then
            getProyectosDeUsuarios.Add m_Proyecto.IDProyecto, m_Proyecto
        End If
siguiente:
        Set m_Proyecto = Nothing
        
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getProyectosDeUsuarios ha devuelto el error: " & Err.description
    End If
End Function
Public Function getProyectosParaTareas( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjProyecto As Proyecto
    
    
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        m_SQL = "SELECT * " & _
                "FROM TbProyectos " & _
                "WHERE CadenaNombreAutorizados Like'*" & m_ObjUsuarioConectado.Nombre & "*'"
    ElseIf EsCalidad = EnumSiNo.Sí Then
        If m_ObjUsuarioParaTareas Is Nothing Then
             m_SQL = "TbProyectos"
        Else
            m_SQL = "SELECT * " & _
                "FROM TbProyectos " & _
                "WHERE CadenaNombreAutorizados Like'*" & m_ObjUsuarioParaTareas.Nombre & "*'"
           
        End If
    Else
    
        m_SQL = "TbProyectos"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjProyecto = New Proyecto
            For Each m_Campo In m_ObjProyecto.ColCampos
                m_ObjProyecto.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getProyectosParaTareas Is Nothing Then
                Set getProyectosParaTareas = New Scripting.Dictionary
                getProyectosParaTareas.CompareMode = TextCompare
            End If
            If Not getProyectosParaTareas.Exists(m_ObjProyecto.IDProyecto) Then
                getProyectosParaTareas.Add m_ObjProyecto.IDProyecto, m_ObjProyecto
            End If
            Set m_ObjProyecto = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getProyectosParaTareas ha devuelto el error: " & Err.description
    End If
End Function

Public Function getProyectosBusqueda( _
                                        Optional p_Activos As String, _
                                        Optional p_NombreJP As String, _
                                        Optional p_PalabraClave As String, _
                                        Optional p_RespCalidad As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    
    Dim m_ObjProyecto As Proyecto
    Dim m_ID As Variant
    
    
    On Error GoTo errores
    If p_Activos = "" And p_NombreJP = "" And p_PalabraClave = "" _
        And p_RespCalidad = "" Then
        Exit Function
    End If
    t1 = Timer
    If m_ObjEntorno.ColProyectosTotales Is Nothing Then
        Exit Function
    End If
    t2 = Timer
    Debug.Print "m_ObjEntorno.ColProyectosTotales", t2 - t1
    t1 = Timer
    For Each m_ID In m_ObjEntorno.ColProyectosTotales
        Set m_ObjProyecto = m_ObjEntorno.ColProyectosTotales(m_ID)
        If p_Activos <> "" Then
            If p_Activos = "Sí" Then
                If m_ObjProyecto.FechaCierre <> "" Then
                    GoTo siguiente
                End If
            ElseIf p_Activos = "No" Then
                If m_ObjProyecto.FechaCierre = "" Then
                    GoTo siguiente
                End If
            End If
            
        End If
        If p_NombreJP <> "" Then
           If InStr(1, m_ObjProyecto.CadenaNombreAutorizados, p_NombreJP) = 0 Then
                GoTo siguiente
           End If
            
        End If
        If p_PalabraClave <> "" Then
            If m_ObjProyecto.Expediente Is Nothing Then
                GoTo siguiente
            End If
            If InStr(1, m_ObjProyecto.NombreProyecto, p_PalabraClave) = 0 And _
                InStr(1, m_ObjProyecto.Proyecto, p_PalabraClave) = 0 Then
                GoTo siguiente
            End If
        End If
        If p_RespCalidad <> "" Then
           If m_ObjProyecto.NombreUsuarioCalidad <> p_RespCalidad Then
                GoTo siguiente
           End If
            
        End If
        If getProyectosBusqueda Is Nothing Then
            Set getProyectosBusqueda = New Scripting.Dictionary
            getProyectosBusqueda.CompareMode = TextCompare
        End If
        If Not getProyectosBusqueda.Exists(m_ObjProyecto.IDProyecto) Then
            getProyectosBusqueda.Add m_ObjProyecto.IDProyecto, m_ObjProyecto
        End If
        
siguiente:
        Set m_ObjProyecto = Nothing
    Next
    t2 = Timer
    Debug.Print "Fin de Filtrado", t2 - t1
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getProyectosBusqueda ha devuelto el error: " & Err.description
    End If
End Function

Public Function getProyectosBusqueda1( _
                                        Optional p_Activos As String, _
                                        Optional p_NombreJP As String, _
                                        Optional p_PalabraClave As String, _
                                        Optional p_RespCalidad As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    Dim m_SQLAlInicio As String
    Dim m_Where As String
    Dim m_WhereJP As String
    Dim m_WhereCalidad As String
    Dim m_WhereActivos As String
    Dim m_WherePalabraClave As String
    Dim m_Proyecto As Proyecto
    
    
    
    On Error GoTo errores
    If p_Activos = "" And p_NombreJP = "" And p_PalabraClave = "" And p_RespCalidad = "" _
        And p_RespCalidad = "" Then
        Exit Function
    End If
    
    m_SQLAlInicio = "SELECT * " & _
                    "FROM TbProyectos "
                    
    If p_Activos = "Sí" Then
        m_WhereActivos = "FechaCierre Is Null "
    ElseIf p_Activos = "No" Then
        m_WhereActivos = "Not FechaCierre Is Null "
    Else
        m_WhereActivos = "(Not FechaCierre Is Null or FechaCierre Is Null) "
    End If
    If p_NombreJP <> "" Then
        m_WhereJP = "CadenaNombreAutorizados Like '*" & p_NombreJP & "*' "
    Else
        m_WhereJP = "(CadenaNombreAutorizados Like '*' or CadenaNombreAutorizados Is Null) "
    End If
    If p_RespCalidad <> "" Then
        m_WhereCalidad = "NombreUsuarioCalidad ='" & p_RespCalidad & "' "
    Else
        m_WhereCalidad = "(NombreUsuarioCalidad Like '*' or NombreUsuarioCalidad Is Null) "
    End If
    If p_PalabraClave <> "" Then
        m_WherePalabraClave = "((NombreProyecto Like '*" & p_PalabraClave & "*')  or (Proyecto Like '*" & p_PalabraClave & "*')) "
    
    End If
    If p_PalabraClave = "" Then
        m_Where = "WHERE " & _
        m_WhereActivos & " AND " & _
        m_WhereJP & " AND " & _
        m_WhereCalidad & ";"
    Else
        m_Where = "WHERE " & _
        m_WhereActivos & " AND " & _
        m_WhereJP & " AND " & _
        m_WhereCalidad & " AND " & _
        m_WherePalabraClave & ";"
    End If
    m_SQL = m_SQLAlInicio & m_Where
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Proyecto = New Proyecto
            For Each m_Campo In m_Proyecto.ColCampos
                m_Proyecto.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getProyectosBusqueda1 Is Nothing Then
                Set getProyectosBusqueda1 = New Scripting.Dictionary
                getProyectosBusqueda1.CompareMode = TextCompare
            End If
            If Not getProyectosBusqueda1.Exists(m_Proyecto.IDProyecto) Then
                getProyectosBusqueda1.Add m_Proyecto.IDProyecto, m_Proyecto
            End If
            Set m_Proyecto = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getProyectosBusqueda1 ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgosBiblioteca( _
                                        Optional p_Familia As String, _
                                        Optional p_PalabraClave As String, _
                                        Optional p_Activo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgoBiblioteca As RiesgoBiblioteca
    Dim m_WherePalabraClave As String
    Dim m_WhereFamilia As String
    Dim m_WhereActivo As String
    Dim m_Where As String
    Dim m_SQLInicial As String
    
    On Error GoTo errores
    
    If p_Familia <> "" Then
        m_WhereFamilia = "((Familia)='" & p_Familia & "')"
    Else
        m_WhereFamilia = "((Familia) Like '*' Or Familia Is Null)"
    End If
    
    If p_PalabraClave <> "" Then
        m_WherePalabraClave = "((Descripcion) Like '*" & p_PalabraClave & "*')"
    Else
        m_WherePalabraClave = "((Descripcion) Like '*' Or Descripcion Is Null)"
    End If
    If p_Activo = "Sí" Or p_Activo = "No" Then
        m_WhereActivo = "((Activo)='" & p_Activo & "')"
    Else
        m_WhereActivo = "((Activo) Like '*' Or Activo Is Null)"
    End If
    
    m_SQLInicial = "SELECT * " & _
                    "FROM TbBibliotecaRiesgos "
    
    m_Where = "WHERE( " & _
                m_WhereFamilia & " " & _
                "AND " & m_WhereActivo & " " & _
                "AND " & m_WherePalabraClave & ");"
    m_SQL = m_SQLInicial & " " & m_Where
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgoBiblioteca = New RiesgoBiblioteca
            For Each m_Campo In m_ObjRiesgoBiblioteca.ColCampos
                m_ObjRiesgoBiblioteca.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosBiblioteca Is Nothing Then
                Set getRiesgosBiblioteca = New Scripting.Dictionary
                getRiesgosBiblioteca.CompareMode = TextCompare
            End If
            If Not getRiesgosBiblioteca.Exists(m_ObjRiesgoBiblioteca.IDRiesgoTipo) Then
                getRiesgosBiblioteca.Add m_ObjRiesgoBiblioteca.IDRiesgoTipo, m_ObjRiesgoBiblioteca
            End If
            Set m_ObjRiesgoBiblioteca = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosBiblioteca ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgosBibliotecasParaRetipificar( _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgoBiblioteca As RiesgoBiblioteca
    
    
    On Error GoTo errores
    
    
    m_SQL = "SELECT * " & _
            "FROM TbBibliotecaRiesgos " & _
            "WHERE Descripcion Like 'Ninguno de los anteriores*';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgoBiblioteca = New RiesgoBiblioteca
            For Each m_Campo In m_ObjRiesgoBiblioteca.ColCampos
                m_ObjRiesgoBiblioteca.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosBibliotecasParaRetipificar Is Nothing Then
                Set getRiesgosBibliotecasParaRetipificar = New Scripting.Dictionary
                getRiesgosBibliotecasParaRetipificar.CompareMode = TextCompare
            End If
            If Not getRiesgosBibliotecasParaRetipificar.Exists(m_ObjRiesgoBiblioteca.IDRiesgoTipo) Then
                getRiesgosBibliotecasParaRetipificar.Add m_ObjRiesgoBiblioteca.IDRiesgoTipo, m_ObjRiesgoBiblioteca
            End If
            Set m_ObjRiesgoBiblioteca = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosBibliotecasParaRetipificar ha devuelto el error: " & Err.description
    End If
End Function
Public Function AlgunRiesgoDeBibliotecaEnProyectosActivos( _
                                                            p_RiesgoBiblioteca As RiesgoBiblioteca, _
                                                            Optional ByRef p_Error As String _
                                                            ) As EnumSiNo

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    
    
    m_SQL = "SELECT distinct TbProyectosEdiciones.IDEdicion " & _
            "FROM (TbProyectosEdiciones " & _
            "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion) " & _
            "INNER JOIN TbBibliotecaRiesgos ON TbRiesgos.CodRiesgoBiblioteca = TbBibliotecaRiesgos.CODIGO " & _
            "WHERE (((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
            "AND ((TbBibliotecaRiesgos.CODIGO)='" & p_RiesgoBiblioteca.codigo & "'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            AlgunRiesgoDeBibliotecaEnProyectosActivos = EnumSiNo.No
        Else
            AlgunRiesgoDeBibliotecaEnProyectosActivos = EnumSiNo.Sí
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.AlgunRiesgoDeBibliotecaEnProyectosActivos ha devuelto el error: " & Err.description
    End If
End Function
Public Function AlgunRiesgoDeBibliotecaEnProyectos( _
                                                    p_RiesgoBiblioteca As RiesgoBiblioteca, _
                                                    Optional ByRef p_Error As String _
                                                    ) As EnumSiNo

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    
    
    m_SQL = "SELECT distinct TbProyectosEdiciones.IDEdicion " & _
            "FROM (TbProyectosEdiciones " & _
            "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion) " & _
            "INNER JOIN TbBibliotecaRiesgos ON TbRiesgos.CodRiesgoBiblioteca = TbBibliotecaRiesgos.CODIGO " & _
            "WHERE TbBibliotecaRiesgos.CODIGO='" & p_RiesgoBiblioteca.codigo & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            AlgunRiesgoDeBibliotecaEnProyectos = EnumSiNo.No
        Else
            AlgunRiesgoDeBibliotecaEnProyectos = EnumSiNo.Sí
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.AlgunRiesgoDeBibliotecaEnProyectos ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgosActivos( _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
        
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        m_SQL = "SELECT TbRiesgos.* " & _
                "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
                "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
                "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE (((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
                "AND ((TbProyectos.CadenaNombreAutorizados) Like'*" & m_ObjUsuarioConectado.Nombre & "*'));"
    Else
        m_SQL = "SELECT TbRiesgos.* " & _
                "FROM TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE (((TbProyectosEdiciones.FechaPublicacion) Is Null));"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            'If m_ObjRiesgo.IDRiesgo = "1271" Then Stop
            'Debug.Print m_ObjRiesgo.IDRiesgo
            If getRiesgosActivos Is Nothing Then
                Set getRiesgosActivos = New Scripting.Dictionary
                getRiesgosActivos.CompareMode = TextCompare
            End If
            If Not getRiesgosActivos.Exists(m_ObjRiesgo.IDRiesgo) Then
                getRiesgosActivos.Add m_ObjRiesgo.IDRiesgo, m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosActivos ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgos( _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
        
    On Error GoTo errores
    
    m_SQL = "TbRiesgos"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            'If m_ObjRiesgo.IDRiesgo = "1271" Then Stop
            'Debug.Print m_ObjRiesgo.IDRiesgo
            If getRiesgos Is Nothing Then
                Set getRiesgos = New Scripting.Dictionary
                getRiesgos.CompareMode = TextCompare
            End If
            If Not getRiesgos.Exists(m_ObjRiesgo.IDRiesgo) Then
                getRiesgos.Add m_ObjRiesgo.IDRiesgo, m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgos ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgosAceptadosORetirados( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
        
    On Error GoTo errores
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE Not FechaJustificacionAceptacionRiesgo Is Null OR " & _
            "Not FechaJustificacionRetiroRiesgo Is Null;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            'If m_ObjRiesgo.IDRiesgo = "1271" Then Stop
            'Debug.Print m_ObjRiesgo.IDRiesgo
            If getRiesgosAceptadosORetirados Is Nothing Then
                Set getRiesgosAceptadosORetirados = New Scripting.Dictionary
                getRiesgosAceptadosORetirados.CompareMode = TextCompare
            End If
            If Not getRiesgosAceptadosORetirados.Exists(m_ObjRiesgo.IDRiesgo) Then
                getRiesgosAceptadosORetirados.Add m_ObjRiesgo.IDRiesgo, m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosAceptadosORetirados ha devuelto el error: " & Err.description
    End If
End Function



Public Function getRiesgo( _
                            Optional ByRef p_IDRiesgo As String, _
                            Optional ByRef p_IDEdicion As String, _
                            Optional ByRef p_CodigoRiesgo As String, _
                            Optional ByRef p_Error As String, _
                            Optional ByRef db As DAO.Database = Nothing) As riesgo

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_db As DAO.Database
    Dim m_Result As riesgo  ' obs #16137 fix: local var

    On Error GoTo errores
    
    ' Resolver db: usar el inyectado o hacer fallback a getdb()
    If db Is Nothing Then
        Set m_db = getdb(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set m_db = db
    End If
    
    If p_IDRiesgo = "" And (p_IDEdicion = "" Or p_CodigoRiesgo = "") Then
        p_Error = "falta p_IDRiesgo y p_IDEdicion o p_CodigoRiesgo"
        Err.Raise 1000
    End If
    If p_IDRiesgo <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbRiesgos " & _
                "WHERE IDRiesgo=" & p_IDRiesgo & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbRiesgos " & _
                "WHERE IDEdicion=" & p_IDEdicion & " AND CodigoRiesgo='" & p_CodigoRiesgo & "';"
    End If
    Set rcdDatos = m_db.OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set m_Result = New riesgo  ' obs #16137 fix: local var
        For Each m_Campo In m_Result.ColCampos
            m_Result.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set getRiesgo = m_Result
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getRiesgo ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getRiesgobyCodUnico( _
                                    Optional ByRef p_IDProyecto As String, _
                                    Optional ByRef p_CodigoUnico As String, _
                                    Optional ByRef p_Error As String _
                                    ) As riesgo
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDProyecto = "" And p_CodigoUnico = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbRiesgos.* " & _
            "FROM TbProyectosEdiciones INNER JOIN TbRiesgos " & _
            "ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
            "WHERE (((TbProyectosEdiciones.IDProyecto)=" & p_IDProyecto & ") " & _
            "AND ((TbRiesgos.CodigoUnico)='" & p_CodigoUnico & "')) " & _
            "ORDER BY TbRiesgos.IDRiesgo;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getRiesgobyCodUnico = New riesgo
        For Each m_Campo In getRiesgobyCodUnico.ColCampos
            getRiesgobyCodUnico.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getRiesgobyCodUnico ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getRiesgoAceptadoEnNacimiento( _
                                                ByRef p_Riesgo As riesgo, _
                                                Optional ByRef p_Error As String _
                                                ) As riesgo
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_Riesgo Is Nothing Then
        Exit Function
    End If
    m_SQL = "SELECT TbRiesgos.* " & _
            "FROM TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
            "WHERE TbRiesgos.CodigoRiesgo='" & p_Riesgo.CodigoRiesgo & "' " & _
            "AND Not TbRiesgos.FechaJustificacionAceptacionRiesgo Is Null " & _
            "AND TbProyectosEdiciones.IDProyecto=" & p_Riesgo.Edicion.IDProyecto & " " & _
            "ORDER BY TbRiesgos.IDEdicion;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getRiesgoAceptadoEnNacimiento = New riesgo
        For Each m_Campo In getRiesgoAceptadoEnNacimiento.ColCampos
            getRiesgoAceptadoEnNacimiento.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getRiesgoAceptadoEnNacimiento ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getRiesgoRetiradoEnNacimiento( _
                                                ByRef p_Riesgo As riesgo, _
                                                Optional ByRef p_Error As String _
                                                ) As riesgo
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_Riesgo Is Nothing Then
        Exit Function
    End If
    m_SQL = "SELECT TbRiesgos.* " & _
            "FROM TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
            "WHERE ((Not (TbRiesgos.FechaRetirado) Is Null) " & _
            "AND ((TbProyectosEdiciones.IDProyecto)=" & p_Riesgo.Edicion.IDProyecto & ") " & _
            "AND ((TbRiesgos.IDEdicion)=" & p_Riesgo.IDEdicion & ") " & _
            " AND ((TbRiesgos.CodigoRiesgo)='" & p_Riesgo.CodigoRiesgo & "')) " & _
            "ORDER BY TbRiesgos.IDEdicion DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getRiesgoRetiradoEnNacimiento = New riesgo
        For Each m_Campo In getRiesgoRetiradoEnNacimiento.ColCampos
            getRiesgoRetiradoEnNacimiento.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getRiesgoRetiradoEnNacimiento ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function



Public Function getPM( _
                        p_IDPM As String, _
                        Optional ByRef p_Error As String _
                        ) As pm
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_IDPM = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosPlanMitigacionPpal " & _
            "WHERE IDMitigacion=" & p_IDPM & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getPM = New pm
        For Each m_Campo In getPM.ColCampos
            getPM.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getPM ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPMUltimo( _
                            p_IDRiesgo As String, _
                            Optional ByRef p_Error As String _
                            ) As pm
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_IDRiesgo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosPlanMitigacionPpal " & _
            "WHERE IDRiesgo=" & p_IDRiesgo & " ORDER BY IDMitigacion Desc;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getPMUltimo = New pm
        For Each m_Campo In getPMUltimo.ColCampos
            getPMUltimo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getPMUltimo ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPMs( _
                        p_IDRiesgo As String, _
                        Optional ByRef p_ParaLista As EnumSiNo, _
                        Optional ByRef p_Error As String _
                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_PM As pm
    On Error GoTo errores
    
    If p_IDRiesgo = "" Then
        Exit Function
    End If
    If p_ParaLista = Empty Then
        p_ParaLista = EnumSiNo.No
    End If
    If p_ParaLista = EnumSiNo.No Then
    
        m_SQL = "SELECT * " & _
                "FROM TbRiesgosPlanMitigacionPpal " & _
                "WHERE IDRiesgo=" & p_IDRiesgo & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbRiesgosPlanMitigacionPpal " & _
                "WHERE IDRiesgo=" & p_IDRiesgo & " ORDER BY TbRiesgosPlanMitigacionPpal.FechaDeActivacion;"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_PM = New pm
            For Each m_Campo In m_PM.ColCampos
                m_PM.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getPMs Is Nothing Then
                Set getPMs = New Scripting.Dictionary
                getPMs.CompareMode = TextCompare
            End If
            If Not getPMs.Exists(m_PM.IDMitigacion) Then
                getPMs.Add m_PM.IDMitigacion, m_PM
            End If
            Set m_PM = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getPMs ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPCs( _
                        p_IDRiesgo As String, _
                        Optional ByRef p_ParaLista As EnumSiNo, _
                        Optional ByRef p_Error As String _
                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_PC As pc
    On Error GoTo errores
    
    If p_IDRiesgo = "" Then
        Exit Function
    End If
    If p_ParaLista = EnumSiNo.No Then
        m_SQL = "SELECT * " & _
                "FROM TbRiesgosPlanContingenciaPpal " & _
                "WHERE IDRiesgo=" & p_IDRiesgo & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbRiesgosPlanContingenciaPpal " & _
                "WHERE IDRiesgo=" & p_IDRiesgo & " ORDER BY TbRiesgosPlanContingenciaPpal.FechaDeActivacion;"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_PC = New pc
            For Each m_Campo In m_PC.ColCampos
                m_PC.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getPCs Is Nothing Then
                Set getPCs = New Scripting.Dictionary
                getPCs.CompareMode = TextCompare
            End If
            If Not getPCs.Exists(m_PC.IDContingencia) Then
                getPCs.Add m_PC.IDContingencia, m_PC
            End If
            Set m_PC = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getPCs ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPC( _
                        p_IDPC As String, _
                        Optional ByRef p_Error As String _
                        ) As pc
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_IDPC = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosPlanContingenciaPpal " & _
            "WHERE IDContingencia=" & p_IDPC & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getPC = New pc
        For Each m_Campo In getPC.ColCampos
            getPC.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getPC ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPCUltimo( _
                            p_IDRiesgo As String, _
                            Optional ByRef p_Error As String _
                            ) As pc
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_IDRiesgo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosPlanContingenciaPpal " & _
            "WHERE IDRiesgo=" & p_IDRiesgo & " ORDER BY IDContingencia desc;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getPCUltimo = New pc
        For Each m_Campo In getPCUltimo.ColCampos
            getPCUltimo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getPCUltimo ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPMAccion( _
                                p_IDPMAccion As String, _
                                Optional ByRef p_Error As String _
                                ) As PMAccion

    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_IDPMAccion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosPlanMitigacionDetalle " & _
            "WHERE IDAccionMitigacion=" & p_IDPMAccion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getPMAccion = New PMAccion
        For Each m_Campo In getPMAccion.ColCampos
            getPMAccion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getPMAccion ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPCAccion( _
                                p_IDPCAccion As String, _
                                Optional ByRef p_Error As String _
                                ) As PCAccion

    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_IDPCAccion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosPlanContingenciaDetalle " & _
            "WHERE IDAccionContingencia=" & p_IDPCAccion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getPCAccion = New PCAccion
        For Each m_Campo In getPCAccion.ColCampos
            getPCAccion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getPCAccion ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPCAcciones( _
                                ByRef p_IDPC As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjPCAccion As PCAccion
    
    On Error GoTo errores
    
    If p_IDPC = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosPlanContingenciaDetalle " & _
            "WHERE IDContingencia=" & p_IDPC & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    Set getPCAcciones = New Scripting.Dictionary
    getPCAcciones.CompareMode = TextCompare
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjPCAccion = New PCAccion
            For Each m_Campo In m_ObjPCAccion.ColCampos
                m_ObjPCAccion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If Not getPCAcciones.Exists(CStr(m_ObjPCAccion.IDAccionContingencia)) Then
                getPCAcciones.Add CStr(m_ObjPCAccion.IDAccionContingencia), m_ObjPCAccion
            End If
            Set m_ObjPCAccion = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getPCAcciones ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPlanesIgualesEnEdcionesAnteriores( _
                                                        p_ObjPlan As Object, _
                                                        Optional ByRef p_Error As String _
                                                        ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjPlan As Object
    Dim m_IDPlan As String
    Dim m_idProyecto As String
    Dim m_CodRiesgo As String
    Dim m_CodPlan As String
    
    'primero he de conocer el IDProyecto,Riesgo,CodPM
    On Error GoTo errores
    
    If Not TypeOf p_ObjPlan Is pm And Not TypeOf p_ObjPlan Is pc Then
        p_Error = "La acción de es un objeto correcto"
        Err.Raise 1000
    End If
    
    If TypeOf p_ObjPlan Is pm Then
        m_IDPlan = p_ObjPlan.IDMitigacion
        m_CodPlan = p_ObjPlan.CodMitigacion
        m_SQL = "SELECT TbProyectosEdiciones.IDProyecto, TbRiesgos.CodigoRiesgo, TbRiesgosPlanMitigacionPpal.CodMitigacion " & _
                "FROM TbProyectosEdiciones INNER JOIN (TbRiesgos INNER JOIN TbRiesgosPlanMitigacionPpal ON " & _
                "TbRiesgos.IDRiesgo = TbRiesgosPlanMitigacionPpal.IDRiesgo) ON " & _
                "TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE (((TbRiesgosPlanMitigacionPpal.IDMitigacion)=" & m_IDPlan & "));"
    Else
        m_IDPlan = p_ObjPlan.IDContingencia
        m_CodPlan = p_ObjPlan.CodContingencia
        m_SQL = "SELECT TbRiesgosPlanContingenciaPpal.CodContingencia, TbRiesgos.CodigoRiesgo, TbProyectosEdiciones.IDProyecto " & _
                "FROM ((TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion) " & _
                "INNER JOIN TbRiesgosPlanContingenciaPpal ON TbRiesgos.IDRiesgo = TbRiesgosPlanContingenciaPpal.IDRiesgo) " & _
                "INNER JOIN TbRiesgosPlanContingenciaDetalle ON " & _
                "TbRiesgosPlanContingenciaPpal.IDContingencia = TbRiesgosPlanContingenciaDetalle.IDContingencia " & _
                "WHERE (((TbRiesgosPlanContingenciaDetalle.IDAccionContingencia)=" & m_IDPlan & "));"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        m_idProyecto = Nz(.fields("IDProyecto"), "")
        If TypeOf p_ObjPlan Is pm Then
            m_CodPlan = Nz(.fields("CodMitigacion"), "")
        Else
            m_CodPlan = Nz(.fields("CodContingencia"), "")
        End If
        m_CodRiesgo = Nz(.fields("CodigoRiesgo"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If m_idProyecto = "" Or m_CodPlan = "" Or m_CodRiesgo = "" Then
        Exit Function
    End If
    If TypeOf p_ObjPlan Is pm Then
        m_SQL = "SELECT TbRiesgosPlanMitigacionPpal.* " & _
                "FROM TbProyectosEdiciones INNER JOIN (TbRiesgos INNER JOIN TbRiesgosPlanMitigacionPpal ON " & _
                "TbRiesgos.IDRiesgo = TbRiesgosPlanMitigacionPpal.IDRiesgo) ON " & _
                "TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE (((TbProyectosEdiciones.IDProyecto)=" & m_idProyecto & _
                ") AND ((TbRiesgos.CodigoRiesgo)='" & m_CodRiesgo & _
                "') AND ((TbRiesgosPlanMitigacionPpal.CodMitigacion)='" & m_CodPlan & _
                "') AND ((TbRiesgosPlanMitigacionPpal.IDMitigacion)<=" & m_IDPlan & "));"
    Else
        m_SQL = "SELECT TbRiesgosPlanContingenciaPpal.* " & _
                "FROM TbProyectosEdiciones INNER JOIN (TbRiesgos INNER JOIN TbRiesgosPlanContingenciaPpal ON " & _
                "TbRiesgos.IDRiesgo = TbRiesgosPlanContingenciaPpal.IDRiesgo) ON " & _
                "TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                "WHERE (((TbProyectosEdiciones.IDProyecto)=" & m_idProyecto & _
                ") AND ((TbRiesgos.CodigoRiesgo)='" & m_CodRiesgo & _
                "') AND ((TbRiesgosPlanContingenciaPpal.CodContingencia)='" & m_CodPlan & _
                "') AND ((TbRiesgosPlanContingenciaPpal.IDContingencia)<=" & m_IDPlan & "));"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        If TypeOf p_ObjPlan Is pm Then
            Set m_ObjPlan = New pm
        Else
            Set m_ObjPlan = New pc
        End If
        
        For Each m_Campo In m_ObjPlan.ColCampos
            m_ObjPlan.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
        If getPlanesIgualesEnEdcionesAnteriores Is Nothing Then
            Set getPlanesIgualesEnEdcionesAnteriores = New Scripting.Dictionary
            getPlanesIgualesEnEdcionesAnteriores.CompareMode = TextCompare
        End If
        If TypeOf p_ObjPlan Is pm Then
            If Not getPlanesIgualesEnEdcionesAnteriores.Exists(CStr(m_ObjPlan.IDMitigacion)) Then
                getPlanesIgualesEnEdcionesAnteriores.Add CStr(m_ObjPlan.IDMitigacion), m_ObjPlan
            End If
        Else
            If Not getPlanesIgualesEnEdcionesAnteriores.Exists(CStr(m_ObjPlan.IDMitigacion)) Then
                getPlanesIgualesEnEdcionesAnteriores.Add CStr(m_ObjPlan.IDMitigacion), m_ObjPlan
            End If
        End If
        Set m_ObjPlan = Nothing
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getPlanesIgualesEnEdcionesAnteriores ha devuelto el error: " & Err.description
    End If
End Function

Public Function getPMAcciones( _
                                p_IDPM As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary

    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjPMAccion As PMAccion
    
    On Error GoTo errores
    
    If p_IDPM = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosPlanMitigacionDetalle " & _
            "WHERE IDMitigacion=" & p_IDPM & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    Set getPMAcciones = New Scripting.Dictionary
    getPMAcciones.CompareMode = TextCompare
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjPMAccion = New PMAccion
            For Each m_Campo In m_ObjPMAccion.ColCampos
                m_ObjPMAccion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If Not getPMAcciones.Exists(CStr(m_ObjPMAccion.IDAccionMitigacion)) Then
                getPMAcciones.Add CStr(m_ObjPMAccion.IDAccionMitigacion), m_ObjPMAccion
            End If
            Set m_ObjPMAccion = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getPMAcciones ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getAnexo( _
                            p_IDAnexo As String, _
                            Optional ByRef p_Error As String _
                            ) As Anexo

    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_IDAnexo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbAnexos " & _
            "WHERE IDAnexo=" & p_IDAnexo & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            
            Exit Function
        End If
        Set getAnexo = New Anexo
        For Each m_Campo In getAnexo.ColCampos
            getAnexo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getAnexo ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getAnexoEvidenciaUTE( _
                                        p_IDEdicion As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Anexo

    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_IDEdicion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbAnexos " & _
            "WHERE IDEdicion=" & p_IDEdicion & " " & _
            "AND EvidenciaUTE='Sí' " & _
            "ORDER BY IDAnexo DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            
            Exit Function
        End If
        Set getAnexoEvidenciaUTE = New Anexo
        For Each m_Campo In getAnexoEvidenciaUTE.ColCampos
            getAnexoEvidenciaUTE.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getAnexoEvidenciaUTE ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getExpediente( _
                                p_IDExpediente As String, _
                                Optional ByRef p_Error As String _
                                ) As Expediente
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    
    
    On Error GoTo errores
    
    If p_IDExpediente = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbExpedientes " & _
            "WHERE IDExpediente=" & p_IDExpediente & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getExpediente = New Expediente
        For Each m_Campo In getExpediente.ColCampos
            'If CStr(m_Campo) = "TipoInforme" Then Stop
            getExpediente.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método getExpediente ha devuelto el error: " & Err.description
    End If
End Function
Public Function getExpedienteResponsables( _
                                            p_IDExpediente As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ExpedienteResponsable As ExpedienteResponsable
    
    
    On Error GoTo errores
    
    If p_IDExpediente = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbExpedientesResponsables " & _
            "WHERE IDExpediente=" & p_IDExpediente & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ExpedienteResponsable = New ExpedienteResponsable
            For Each m_Campo In m_ExpedienteResponsable.ColCampos
                m_ExpedienteResponsable.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                 If p_Error <> "" Then
                     Err.Raise 1000
                 End If
             Next
             If getExpedienteResponsables Is Nothing Then
                Set getExpedienteResponsables = New Scripting.Dictionary
                getExpedienteResponsables.CompareMode = TextCompare
             End If
             If Not getExpedienteResponsables.Exists(CStr(m_ExpedienteResponsable.IDExpedienteResponsable)) Then
                getExpedienteResponsables.Add CStr(m_ExpedienteResponsable.IDExpedienteResponsable), m_ExpedienteResponsable
             End If
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getExpedienteResponsables ha devuelto el error: " & Err.description
    End If
End Function
Public Function getExpedienteEntidad( _
                                        p_IDExpediente As String, _
                                        Optional ByRef p_Error As String _
                                        ) As ExpedienteEntidad
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    
    
    On Error GoTo errores
    
    If p_IDExpediente = "" Then
        Exit Function
    End If
     m_SQL = "SELECT * " & _
            "FROM TbExpedientesConEntidades " & _
            "WHERE IDExpediente=" & p_IDExpediente & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getExpedienteEntidad = New ExpedienteEntidad
        For Each m_Campo In getExpedienteEntidad.ColCampos
            getExpedienteEntidad.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método getExpedienteEntidad ha devuelto el error: " & Err.description
    End If
End Function

Public Function getRiesgoExterno( _
                                    p_IDRiesgoExt As String, _
                                    Optional ByRef p_Error As String _
                                    ) As RiesgoExterno
                            
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_IDRiesgoExt = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosAIntegrar " & _
            "WHERE IDRiesgoExt=" & p_IDRiesgoExt & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getRiesgoExterno = New RiesgoExterno
        For Each m_Campo In getRiesgoExterno.ColCampos
            getRiesgoExterno.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getRiesgoExterno ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getRiesgoBiblioteca( _
                                    Optional p_IDRiesgoBiblioteca As String, _
                                    Optional p_CodRiesgo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As RiesgoBiblioteca
                            
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_IDRiesgoBiblioteca = "" And p_CodRiesgo = "" Then
        Exit Function
    End If
    If p_IDRiesgoBiblioteca <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbBibliotecaRiesgos " & _
                "WHERE IDRiesgoTipo=" & p_IDRiesgoBiblioteca & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbBibliotecaRiesgos " & _
                "WHERE CODIGO='" & p_CodRiesgo & "';"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getRiesgoBiblioteca = New RiesgoBiblioteca
        For Each m_Campo In getRiesgoBiblioteca.ColCampos
            getRiesgoBiblioteca.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getRiesgoBiblioteca ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getRiesgoBibliotecaNoExiste( _
                                            Optional ByRef p_Error As String _
                                                ) As RiesgoBiblioteca
                            
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    m_SQL = "SELECT * " & _
                "FROM TbBibliotecaRiesgos " & _
                "WHERE Descripcion Like 'Ninguno de los anteriores*';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getRiesgoBibliotecaNoExiste = New RiesgoBiblioteca
        For Each m_Campo In getRiesgoBibliotecaNoExiste.ColCampos
            getRiesgoBibliotecaNoExiste.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getRiesgoBibliotecaNoExiste ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getUsuario( _
                            Optional p_Id As String, _
                            Optional p_UsuarioRed As String, _
                            Optional p_Nombre As String, _
                            Optional p_Correo As String, _
                            Optional ByRef p_Error As String _
                            ) As Usuario

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_NombreCampoID As String
    Dim m_EsNumeroID As Boolean
    Dim m_Where As String
    Dim m_ValorID As String
    Dim m_SQLInicial As String
    Dim m_ObjUsuario As Usuario
    
    On Error GoTo errores
    
    If p_Id = "" And p_UsuarioRed = "" And p_Nombre = "" And p_Correo = "" Then
        p_Error = "falta p_ID y p_UsuarioRed y p_Nombre y p_Correo"
        Err.Raise 1000
    End If
    m_SQLInicial = "SELECT TbUsuariosAplicaciones.* " & _
                    "FROM TbUsuariosAplicaciones "
    If p_Id <> "" Then
        m_NombreCampoID = "ID"
        m_ValorID = p_Id
        m_EsNumeroID = True
    ElseIf p_UsuarioRed <> "" Then
        m_NombreCampoID = "UsuarioRed"
        m_ValorID = p_UsuarioRed
        m_EsNumeroID = False
    ElseIf p_Nombre <> "" Then
        m_NombreCampoID = "Nombre"
        m_ValorID = p_Nombre
        m_EsNumeroID = False
    ElseIf p_Correo <> "" Then
        m_NombreCampoID = "CorreoUsuario"
        m_ValorID = p_Correo
        m_EsNumeroID = False
    End If
    If m_EsNumeroID Then
        m_Where = m_NombreCampoID & "=" & m_ValorID & ";"
    Else
        m_Where = m_NombreCampoID & "='" & m_ValorID & "';"
    End If
    m_SQL = m_SQLInicial & "WHERE " & m_Where
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        Set getUsuario = New Usuario
        For Each m_Campo In getUsuario.ColCampos
            getUsuario.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método constructor.getUsuario ha devuelto el error: " & Err.description
    End If
End Function

Public Function getUsuariosAdministradores( _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjUsuario As Usuario
    
    
    On Error GoTo errores
    
    m_SQL = "SELECT * " & _
            "FROM TbUsuariosAplicaciones " & _
            "WHERE EsAdministrador='Sí';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_ObjUsuario = New Usuario
                For Each m_Campo In m_ObjUsuario.ColCampos
                    m_ObjUsuario.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getUsuariosAdministradores Is Nothing Then
                    Set getUsuariosAdministradores = New Scripting.Dictionary
                    getUsuariosAdministradores.CompareMode = TextCompare
                End If
                If Not getUsuariosAdministradores.Exists(m_ObjUsuario.UsuarioRed) Then
                    getUsuariosAdministradores.Add m_ObjUsuario.UsuarioRed, m_ObjUsuario
                End If
                Set m_ObjUsuario = Nothing
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    m_SQL = "SELECT TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones INNER JOIN TbUsuariosAplicacionesPermisos ON " & _
            "TbUsuariosAplicaciones.CorreoUsuario = TbUsuariosAplicacionesPermisos.CorreoUsuario " & _
            "WHERE (((TbUsuariosAplicacionesPermisos.IDAplicacion)=" & m_IDAplicacion & _
            ") AND ((TbUsuariosAplicaciones.[Activado])=True) " & _
            "AND ((TbUsuariosAplicacionesPermisos.EsUsuarioAdministrador)='Sí'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_ObjUsuario = New Usuario
                For Each m_Campo In m_ObjUsuario.ColCampos
                    m_ObjUsuario.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getUsuariosAdministradores Is Nothing Then
                    Set getUsuariosAdministradores = New Scripting.Dictionary
                    getUsuariosAdministradores.CompareMode = TextCompare
                End If
                If Not getUsuariosAdministradores.Exists(CStr(m_ObjUsuario.UsuarioRed)) Then
                    getUsuariosAdministradores.Add CStr(m_ObjUsuario.UsuarioRed), m_ObjUsuario
                End If
                
                Set m_ObjUsuario = Nothing
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getUsuariosAdministradores ha devuelto el error: " & Err.description
    End If
End Function

Public Function getUsuariosCalidad( _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjUsuario As Usuario
    
    
    On Error GoTo errores
    
    m_SQL = "SELECT TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones INNER JOIN TbUsuariosAplicacionesPermisos " & _
            "ON TbUsuariosAplicaciones.CorreoUsuario = TbUsuariosAplicacionesPermisos.CorreoUsuario " & _
            "WHERE " & _
            "FechaBaja Is Null " & _
            "AND EsUsuarioCalidad='Sí' " & _
            "AND TbUsuariosAplicacionesPermisos.IDAplicacion=" & m_IDAplicacion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_ObjUsuario = New Usuario
                For Each m_Campo In m_ObjUsuario.ColCampos
                    m_ObjUsuario.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getUsuariosCalidad Is Nothing Then
                    Set getUsuariosCalidad = New Scripting.Dictionary
                    getUsuariosCalidad.CompareMode = TextCompare
                End If
                If Not getUsuariosCalidad.Exists(CStr(m_ObjUsuario.UsuarioRed)) Then
                    getUsuariosCalidad.Add CStr(m_ObjUsuario.UsuarioRed), m_ObjUsuario
                End If
                
                Set m_ObjUsuario = Nothing
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getUsuariosCalidad ha devuelto el error: " & Err.description
    End If
End Function
Public Function getUsuariosCalidadAvisos( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjUsuario As Usuario
    
    
    On Error GoTo errores
    
    m_SQL = "SELECT TbUsuariosAplicaciones.* " & _
            "FROM TbUsuariosAplicaciones INNER JOIN TbUsuariosAplicacionesPermisos ON " & _
            "TbUsuariosAplicaciones.CorreoUsuario = TbUsuariosAplicacionesPermisos.CorreoUsuario " & _
            "WHERE (((TbUsuariosAplicacionesPermisos.IDAplicacion)=" & m_IDAplicacion & _
            ") AND ((TbUsuariosAplicaciones.[Activado])=True) " & _
            "AND ((TbUsuariosAplicacionesPermisos.EsUsuarioCalidad)='Sí') " & _
            "AND ((TbUsuariosAplicacionesPermisos.EsUsuarioCalidadAvisos)='Sí'));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_ObjUsuario = New Usuario
                For Each m_Campo In m_ObjUsuario.ColCampos
                    m_ObjUsuario.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getUsuariosCalidadAvisos Is Nothing Then
                    Set getUsuariosCalidadAvisos = New Scripting.Dictionary
                    getUsuariosCalidadAvisos.CompareMode = TextCompare
                End If
                If Not getUsuariosCalidadAvisos.Exists(CStr(m_ObjUsuario.UsuarioRed)) Then
                    getUsuariosCalidadAvisos.Add CStr(m_ObjUsuario.UsuarioRed), m_ObjUsuario
                End If
                
                Set m_ObjUsuario = Nothing
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getUsuariosCalidadAvisos ha devuelto el error: " & Err.description
    End If
End Function


Public Function getUsuariosTecnicos( _
                                    Optional p_SoloActivos As EnumSiNo = EnumSiNo.Sí, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjUsuario As Usuario
    
    
    On Error GoTo errores
    If p_SoloActivos = Empty Then
        p_SoloActivos = EnumSiNo.Sí
    End If
    If p_SoloActivos = EnumSiNo.Sí Then
        m_SQL = "SELECT * " & _
                "FROM TbUsuariosAplicaciones " & _
                "WHERE " & _
                "FechaBaja Is Null " & _
                "AND EsAdministrador<>'Sí' " & _
                "ORDER BY TbUsuariosAplicaciones.Nombre;"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbUsuariosAplicaciones " & _
                "WHERE " & _
                "EsAdministrador<>'Sí' " & _
                "ORDER BY TbUsuariosAplicaciones.Nombre;"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_ObjUsuario = New Usuario
                For Each m_Campo In m_ObjUsuario.ColCampos
                    m_ObjUsuario.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getUsuariosTecnicos Is Nothing Then
                    Set getUsuariosTecnicos = New Scripting.Dictionary
                    getUsuariosTecnicos.CompareMode = TextCompare
                End If
                If Not getUsuariosTecnicos.Exists(CStr(m_ObjUsuario.UsuarioRed)) Then
                    getUsuariosTecnicos.Add CStr(m_ObjUsuario.UsuarioRed), m_ObjUsuario
                End If
                
                Set m_ObjUsuario = Nothing
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getUsuariosTecnicos ha devuelto el error: " & Err.description
    End If
End Function

Public Function getUsuariosTecnicosConAlgunaGestion( _
                                                        Optional ByRef p_Error As String _
                                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjUsuario As Usuario
    
    
    On Error GoTo errores
   m_SQL = "SELECT distinct TbUsuariosAplicaciones.* " & _
            "FROM (TbProyectos INNER JOIN TbExpedientesResponsables " & _
            "ON TbProyectos.IDExpediente = TbExpedientesResponsables.IdExpediente) " & _
            "INNER JOIN TbUsuariosAplicaciones ON TbExpedientesResponsables.IdUsuario = TbUsuariosAplicaciones.Id " & _
            "ORDER BY TbUsuariosAplicaciones.Nombre;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_ObjUsuario = New Usuario
                For Each m_Campo In m_ObjUsuario.ColCampos
                    m_ObjUsuario.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getUsuariosTecnicosConAlgunaGestion Is Nothing Then
                    Set getUsuariosTecnicosConAlgunaGestion = New Scripting.Dictionary
                    getUsuariosTecnicosConAlgunaGestion.CompareMode = TextCompare
                End If
                If Not getUsuariosTecnicosConAlgunaGestion.Exists(CStr(m_ObjUsuario.UsuarioRed)) Then
                    getUsuariosTecnicosConAlgunaGestion.Add CStr(m_ObjUsuario.UsuarioRed), m_ObjUsuario
                End If
                
                Set m_ObjUsuario = Nothing
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.ColUsuariosTecnicosConAlgunaGestion ha devuelto el error: " & Err.description
    End If
End Function

Public Function getColJPs( _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjUsuario As Usuario
    
    
    On Error GoTo errores
    
    m_SQL = "SELECT distinct TbUsuariosAplicaciones.* " & _
            "FROM ((TbProyectos INNER JOIN TbExpedientes1 ON TbProyectos.IDExpediente = TbExpedientes1.IDExpediente) " & _
            "INNER JOIN TbExpedientesResponsables ON TbExpedientes1.IDExpediente = TbExpedientesResponsables.IdExpediente) " & _
            "INNER JOIN TbUsuariosAplicaciones ON TbExpedientesResponsables.IdUsuario = TbUsuariosAplicaciones.Id;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_ObjUsuario = New Usuario
                For Each m_Campo In m_ObjUsuario.ColCampos
                    m_ObjUsuario.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getColJPs Is Nothing Then
                    Set getColJPs = New Scripting.Dictionary
                    getColJPs.CompareMode = TextCompare
                End If
                If Not getColJPs.Exists(CStr(m_ObjUsuario.UsuarioRed)) Then
                    getColJPs.Add CStr(m_ObjUsuario.UsuarioRed), m_ObjUsuario
                End If
                
                Set m_ObjUsuario = Nothing
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getColJPs ha devuelto el error: " & Err.description
    End If
End Function

Public Function getAplicacionesPermisos( _
                                            p_CorreoUsuario As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjUsuarioAplicacionPermisos As UsuarioAplicacionPermisos

    On Error GoTo errores

    If p_CorreoUsuario = "" Then
        p_Error = "falta p_CorreoUsuario"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbUsuariosAplicacionesPermisos.* " & _
            "FROM TbUsuariosAplicacionesPermisos " & _
            "WHERE CorreoUsuario='" & p_CorreoUsuario & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjUsuarioAplicacionPermisos = New UsuarioAplicacionPermisos
            For Each m_Campo In m_ObjUsuarioAplicacionPermisos.ColCampos
                m_ObjUsuarioAplicacionPermisos.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getAplicacionesPermisos Is Nothing Then
                Set getAplicacionesPermisos = New Scripting.Dictionary
                getAplicacionesPermisos.CompareMode = TextCompare
            End If
            If Not getAplicacionesPermisos.Exists(CStr(m_ObjUsuarioAplicacionPermisos.IDAplicacion)) Then
                getAplicacionesPermisos.Add m_ObjUsuarioAplicacionPermisos.IDAplicacion, m_ObjUsuarioAplicacionPermisos
            End If
            Set m_ObjUsuarioAplicacionPermisos = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getAplicacionesPermisos ha devuelto el error: " & Err.description
    End If
End Function

Public Function getDocumentoByID( _
                                    p_IDDocumento As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Documento

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_IDDocumento = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
                "FROM TbDocumentos " & _
                "WHERE IDDocumento=" & p_IDDocumento & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getDocumentoByID = New Documento
        For Each m_Campo In getDocumentoByID.ColCampos
            getDocumentoByID.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getDocumentoByID ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getDocumentoByCodigo( _
                                        p_Codigo As String, _
                                        p_Edicion As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Documento

    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
       
    On Error GoTo errores
    
    If p_Codigo = "" Or p_Edicion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
                "FROM TbDocumentos " & _
                "WHERE Codigo='" & p_Codigo & "' AND EDICION=" & p_Edicion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getDocumentoByCodigo = New Documento
        For Each m_Campo In getDocumentoByCodigo.ColCampos
            getDocumentoByCodigo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getDocumentoByCodigo ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getDocumentosbyCodigo( _
                                        p_Codigo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjDocumento As Documento
    
    On Error GoTo errores
    
    If p_Codigo = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbDocumentos " & _
            "WHERE Codigo='" & p_Codigo & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjDocumento = New Documento
            For Each m_Campo In m_ObjDocumento.ColCampos
                m_ObjDocumento.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getDocumentosbyCodigo Is Nothing Then
                Set getDocumentosbyCodigo = New Scripting.Dictionary
                getDocumentosbyCodigo.CompareMode = TextCompare
            End If
            If Not getDocumentosbyCodigo.Exists(CStr(m_ObjDocumento.IDDocumento)) Then
                getDocumentosbyCodigo.Add m_ObjDocumento.IDDocumento, m_ObjDocumento
            End If
            Set m_ObjDocumento = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getDocumentosbyCodigo ha devuelto el error: " & Err.description
    End If
End Function


Public Function getRiesgoByPriorizacion( _
                                        p_IDEdicion As String, _
                                        p_Priorizacion As Single, _
                                        Optional ByRef p_Error As String _
                                        ) As riesgo
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDEdicion = "" Or p_Priorizacion = 0 Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDEdicion=" & p_IDEdicion & "AND Priorizacion=" & p_Priorizacion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getRiesgoByPriorizacion = New riesgo
        For Each m_Campo In getRiesgoByPriorizacion.ColCampos
            getRiesgoByPriorizacion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getRiesgoByPriorizacion ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function







Public Function getNC( _
                        m_IDNC As String, _
                        Optional ByRef p_Error As String _
                        ) As NC

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
        
    On Error GoTo errores
    
    If m_IDNC = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbNoConformidades " & _
            "WHERE IDNoConformidad=" & m_IDNC & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getNC = New NC
        For Each m_Campo In getNC.ColCampos
            getNC.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método constructor.getNC ha devuelto el error: " & Err.description
    End If
End Function



Public Function getValoresDistintos( _
                                    p_NombreTabla As String, _
                                    p_NombreCampo As String, _
                                    Optional p_db As DAO.Database, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_valor As String
    
    On Error GoTo errores
    
    If p_db Is Nothing Then
        Set p_db = getdb()
    End If
    If p_NombreTabla = "" Or p_NombreCampo = "" Then
        p_Error = "falta p_NombreTabla o p_NombreCampo"
        Err.Raise 1000
    End If
    m_SQL = "SELECT DISTINCT " & p_NombreTabla & "." & p_NombreCampo & " " & _
            "From " & p_NombreTabla & " " & _
            "WHERE ((Not (" & p_NombreTabla & "." & p_NombreCampo & ") Is Null));"
    Set rcdDatos = p_db.OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_valor = .fields(p_NombreCampo)
            If getValoresDistintos Is Nothing Then
                Set getValoresDistintos = New Scripting.Dictionary
                getValoresDistintos.CompareMode = TextCompare
            End If
            If Not getValoresDistintos.Exists(m_valor) Then
                getValoresDistintos.Add m_valor, m_valor
            End If
            .MoveNext
            Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getValoresDistintos ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getListaExplicaciones( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ValorNodo As String
    Dim m_ValorTitulo As String
    Dim m_ValorExplicacion As String
    
    On Error GoTo errores
    
    m_SQL = "TbTareasExplicaciones"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_ValorNodo = .fields("NodoTarea")
           ' If m_ValorNodo = "RIESGOSRETIPIFICADOS" Then Stop
            m_ValorTitulo = .fields("TituloTarea")
            m_ValorExplicacion = .fields("Explicacion")
            If getListaExplicaciones Is Nothing Then
                Set getListaExplicaciones = New Scripting.Dictionary
                getListaExplicaciones.CompareMode = TextCompare
            End If
            If Not getListaExplicaciones.Exists(m_ValorNodo) Then
                getListaExplicaciones.Add m_ValorNodo, m_ValorTitulo & "|" & m_ValorExplicacion
            End If
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getListaExplicaciones ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getRiesgoEdicionActiva( _
                                        p_ObjRiesgo As riesgo, _
                                        Optional ByRef p_Error As String _
                                        ) As riesgo

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim NumeroRegistros As Long
    Dim m_IDEdicionActiva As String
    Dim m_CodUnico As String
    
    On Error GoTo errores
    
    If p_ObjRiesgo Is Nothing Then
        Exit Function
    End If
    m_CodUnico = p_ObjRiesgo.CodigoUnico
    If p_ObjRiesgo.Edicion.Proyecto.EdicionUltima Is Nothing Then
        Exit Function
    End If
    m_IDEdicionActiva = p_ObjRiesgo.Edicion.Proyecto.EdicionUltima.IDEdicion
    m_SQL = "SELECT TbRiesgos.* " & _
            "FROM TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
            "WHERE (((TbRiesgos.CodigoUnico)='" & m_CodUnico & _
            "') AND ((TbProyectosEdiciones.IDEdicion)=" & m_IDEdicionActiva & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getRiesgoEdicionActiva = New riesgo
        For Each m_Campo In getRiesgoEdicionActiva.ColCampos
            getRiesgoEdicionActiva.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método constructor.getRiesgoEdicionActiva ha devuelto el error: " & Err.description
    End If
End Function

Public Function getValoracion( _
                                ByRef p_ImpactoGlobal As String, _
                                ByRef p_Vulnerabilidad As String, _
                                Optional ByRef p_Error As String _
                                ) As String

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    If p_Vulnerabilidad = "" Or p_ImpactoGlobal = "" Then
        p_Error = "falta p_Vulnerabilidad o p_ImpactoGlobal"
        Err.Raise 1000
    End If
    m_SQL = "SELECT Valoracion " & _
            "FROM TbRiesgosValoracion " & _
            "WHERE Impacto='" & p_ImpactoGlobal & _
                    "' AND Vulnerabilidad='" & p_Vulnerabilidad & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        getValoracion = Nz(.fields("Valoracion"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getValoracion ha devuelto el error: " & Err.description
    End If
End Function
Public Function getEstadosRiesgoConMaterializaciones( _
                                                    p_ObjRiesgo As riesgo, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Estado As String
    Dim m_FechaEstado As String
    Dim m_Riesgo As riesgo
    
    On Error GoTo errores
    If p_ObjRiesgo Is Nothing Then
        p_Error = "Se ha de indicar El Riesgo"
        Exit Function
    End If
    
    If p_ObjRiesgo.CodigoUnico = "" Then
        p_Error = "Se ha de indicar m_CodigoUnico"
        Exit Function
    End If
    m_SQL = "SELECT TbRiesgos.* " & _
            "FROM TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
            "WHERE (((TbRiesgos.CodigoUnico)='" & p_ObjRiesgo.CodigoUnico & _
            "') AND ((TbProyectosEdiciones.IDProyecto)=" & p_ObjRiesgo.Edicion.IDProyecto & _
            ") AND ((TbProyectosEdiciones.IDEdicion)<=" & p_ObjRiesgo.IDEdicion & "));"
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
                m_Riesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            m_Estado = "Detectado"
            m_FechaEstado = m_Riesgo.FechaDetectadoCalculado
            If getEstadosRiesgoConMaterializaciones Is Nothing Then
                Set getEstadosRiesgoConMaterializaciones = New Scripting.Dictionary
                getEstadosRiesgoConMaterializaciones.CompareMode = TextCompare
            End If
            If Not getEstadosRiesgoConMaterializaciones.Exists(m_Estado) Then
                getEstadosRiesgoConMaterializaciones.Add m_Estado, m_FechaEstado
            End If
            m_Estado = m_Riesgo.ESTADOCalculadoTexto
            m_FechaEstado = m_Riesgo.FechaEstado
            If getEstadosRiesgoConMaterializaciones Is Nothing Then
                Set getEstadosRiesgoConMaterializaciones = New Scripting.Dictionary
                getEstadosRiesgoConMaterializaciones.CompareMode = TextCompare
            End If
            If Not getEstadosRiesgoConMaterializaciones.Exists(m_Estado) Then
                getEstadosRiesgoConMaterializaciones.Add m_Estado, m_FechaEstado
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
        p_Error = "El método constructor.getEstadosRiesgoConMaterializaciones ha devuelto el error: " & Err.description
    End If
End Function
Public Function getEstadosRiesgo( _
                                    p_ObjRiesgo As riesgo, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Estado As String
    Dim m_FechaEstado As String
    Dim m_Riesgo As riesgo
    
    On Error GoTo errores
    If p_ObjRiesgo Is Nothing Then
        p_Error = "Se ha de indicar El Riesgo"
        Exit Function
    End If
    
    If p_ObjRiesgo.CodigoUnico = "" Then
        p_Error = "Se ha de indicar m_CodigoUnico"
        Exit Function
    End If
    m_SQL = "SELECT TbRiesgos.* " & _
            "FROM TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
            "WHERE (((TbRiesgos.CodigoUnico)='" & p_ObjRiesgo.CodigoUnico & _
            "') AND ((TbProyectosEdiciones.IDProyecto)=" & p_ObjRiesgo.Edicion.IDProyecto & _
            ") AND ((TbProyectosEdiciones.IDEdicion)<=" & p_ObjRiesgo.IDEdicion & "));"
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
                m_Riesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            m_Estado = "Detectado"
            m_FechaEstado = m_Riesgo.FechaDetectadoCalculado
            If getEstadosRiesgo Is Nothing Then
                Set getEstadosRiesgo = New Scripting.Dictionary
                getEstadosRiesgo.CompareMode = TextCompare
            End If
            If Not getEstadosRiesgo.Exists(m_Estado) Then
                getEstadosRiesgo.Add m_Estado, m_FechaEstado
            End If
            m_Estado = m_Riesgo.ESTADOCalculadoTexto
            m_FechaEstado = m_Riesgo.FechaEstado
            If getEstadosRiesgo Is Nothing Then
                Set getEstadosRiesgo = New Scripting.Dictionary
                getEstadosRiesgo.CompareMode = TextCompare
            End If
            If Not getEstadosRiesgo.Exists(m_Estado) Then
                getEstadosRiesgo.Add m_Estado, m_FechaEstado
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
        p_Error = "El método constructor.getEstadosRiesgo ha devuelto el error: " & Err.description
    End If
End Function

Public Function getEdicionEnLaQueNaceElRiesgo( _
                                                ByRef p_ObjRiesgo As riesgo, _
                                                Optional ByRef p_Error As String _
                                                ) As Edicion

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_SQLLimitante As String
    
    On Error GoTo errores
    
    If p_ObjRiesgo Is Nothing Then
        Exit Function
    End If
    If p_ObjRiesgo.CodigoUnico = "" Then
        p_Error = "Se ha de indicar CodigoUnico"
        Exit Function
    End If
    m_SQLLimitante = "SELECT Min(TbRiesgos.IDEdicion) AS MínDeIDEdicion " & _
                    "FROM TbRiesgos " & _
                    "WHERE CodigoUnico='" & p_ObjRiesgo.CodigoUnico & "';"
    m_SQL = "SELECT * " & _
            "FROM TbProyectosEdiciones " & _
            "WHERE IDEdicion In (" & m_SQLLimitante & ");"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getEdicionEnLaQueNaceElRiesgo = New Edicion
        For Each m_Campo In getEdicionEnLaQueNaceElRiesgo.ColCampos
            getEdicionEnLaQueNaceElRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método constructor.getEdicionEnLaQueNaceElRiesgo ha devuelto el error: " & Err.description
    End If
End Function

Public Function getEdicionUltimaEnLaQueExiste( _
                                                ByRef p_ObjRiesgo As riesgo, _
                                                Optional ByRef p_Error As String _
                                                ) As Edicion

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_SQLLimitante As String
    Dim m_idProyecto As String
    Dim m_CodRiesgo As String
    
    On Error GoTo errores
    
    If p_ObjRiesgo Is Nothing Then
        Exit Function
    End If
    
    m_idProyecto = p_ObjRiesgo.Edicion.IDProyecto
    m_SQLLimitante = "SELECT Max(TbRiesgos.IDEdicion) AS MáxDeIDEdicion " & _
                    "FROM TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
                    "WHERE (((TbRiesgos.CodigoRiesgo)='" & p_ObjRiesgo.CodigoRiesgo & "') " & _
                    "AND ((TbProyectosEdiciones.IDProyecto)=" & m_idProyecto & "));"
    m_SQL = "SELECT * " & _
            "FROM TbProyectosEdiciones " & _
            "WHERE IDEdicion In (" & m_SQLLimitante & ");"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getEdicionUltimaEnLaQueExiste = New Edicion
        For Each m_Campo In getEdicionUltimaEnLaQueExiste.ColCampos
            getEdicionUltimaEnLaQueExiste.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método constructor.getEdicionUltimaEnLaQueExiste ha devuelto el error: " & Err.description
    End If
End Function

Public Function getRiesgosByCodUnico( _
                                        p_CodigoUnico As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    
    If p_CodigoUnico = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE CodigoUnico='" & p_CodigoUnico & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosByCodUnico Is Nothing Then
                Set getRiesgosByCodUnico = New Scripting.Dictionary
                getRiesgosByCodUnico.CompareMode = TextCompare
            End If
            If Not getRiesgosByCodUnico.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                getRiesgosByCodUnico.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosByCodUnico ha devuelto el error: " & Err.description
    End If
End Function
Public Function getPrimerRiesgoByCodUnico( _
                                            p_CodigoUnico As String, _
                                            Optional ByRef p_Error As String _
                                            ) As riesgo

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_CodigoUnico = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE CodigoUnico='" & p_CodigoUnico & "' " & _
            "ORDER BY IDRiesgo;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getPrimerRiesgoByCodUnico = New riesgo
        For Each m_Campo In getPrimerRiesgoByCodUnico.ColCampos
            getPrimerRiesgoByCodUnico.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método constructor.getPrimerRiesgoByCodUnico ha devuelto el error: " & Err.description
    End If
End Function

'-------------------------------------------
' Nombre: getExpedienteJuridicas
' Propósito: Obtener adjudicatarios del expediente (nodos raíz) según el modelo jerárquico
' Parámetros:
'   - p_IDExpediente As String: identificador del expediente
'   - p_Error As String (ByRef, opcional): detalle del error si ocurre
' Retorno: Scripting.Dictionary (clave: IDSuministrador; valor: objeto Suministrador)
' Dependencias:
'   - getdb() para la conexión a Expedientes_datos.accdb
'   - Clase Suministrador con método ColCampos/SetPropiedad
'-------------------------------------------
Public Function getExpedienteJuridicas( _
                                            p_IDExpediente As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    On Error GoTo ManejoError
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Suministrador As Suministrador
    
    ' Validación de parámetro obligatorio
    If Nz(p_IDExpediente, "") = "" Then
        p_Error = "El parámetro p_IDExpediente es obligatorio."
        Exit Function
    End If
    
    ' Nueva lógica (MAIN): raíz y contratista principal
    m_SQL = "SELECT T.* " & _
            "FROM TbExpedientesSuministradores R " & _
            "INNER JOIN TbSuministradores T ON R.IDSuministrador = T.IDSuministrador " & _
            "WHERE R.IDExpediente=" & p_IDExpediente & " " & _
            "AND (R.IdPadre Is Null AND R.ContratistaPrincipal='Sí');"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL, dbOpenSnapshot)
    With rcdDatos
        If .EOF Then
            ' No hay adjudicatarios raíz para el expediente
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        
        .MoveFirst
        Do While Not .EOF
            Set m_Suministrador = New Suministrador
            For Each m_Campo In m_Suministrador.ColCampos
                m_Suministrador.SetPropiedad m_Campo, Nz(.fields(CStr(m_Campo)).value, ""), p_Error
                If Nz(p_Error, "") <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getExpedienteJuridicas Is Nothing Then
                Set getExpedienteJuridicas = New Scripting.Dictionary
                getExpedienteJuridicas.CompareMode = TextCompare
            End If
            
            If Not getExpedienteJuridicas.Exists(CStr(m_Suministrador.idSuministrador)) Then
                getExpedienteJuridicas.Add CStr(m_Suministrador.idSuministrador), m_Suministrador
            End If
            
            .MoveNext
        Loop
    End With
    
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
ManejoError:
    If Err.Number <> 0 Then
    p_Error = "Error en getExpedienteJuridicas (" & Err.Number & "): " & Err.description
    End If
    
End Function
'-------------------------------------------
' Nombre: GetCadenaJerarquicaEmpresas
' Propósito: Generar cadena de empresas para un expediente según el tipo ("MAIN" adjudicatarios, "SUB" resto)
' Parámetros:
'   - p_IDExpediente As Variant: identificador del expediente (numérico o texto)
'   - sTipo As String: "MAIN" o "SUB"
'   - p_Error As String (ByRef, opcional): detalle del error si ocurre
' Retorno: String (ej. "TDE|INDRA|ACCENTURE")
' Dependencias:
'   - getdb() para la conexión a Expedientes_datos.accdb
'-------------------------------------------
Public Function GetCadenaJerarquicaEmpresas( _
                                            p_IDExpediente As Variant, _
                                            sTipo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    On Error GoTo ManejoError
    Dim rcd As DAO.Recordset
    Dim m_SQL As String
    Dim sResultado As String
    Dim sNombre As String
    
    ' Validación de parámetros
    If Nz(p_IDExpediente, "") = "" Then
        p_Error = "El parámetro p_IDExpediente es obligatorio."
        Exit Function
    End If
    If UCase$(Nz(sTipo, "")) <> "MAIN" And UCase$(Nz(sTipo, "")) <> "SUB" Then
        p_Error = "El parámetro sTipo debe ser 'MAIN' o 'SUB'."
        Exit Function
    End If
    
    ' Construcción de la consulta según el tipo
    m_SQL = "SELECT T.Nemotecnico, T.Nombre " & _
            "FROM TbExpedientesSuministradores R " & _
            "INNER JOIN TbSuministradores T ON R.IDSuministrador = T.IDSuministrador " & _
            "WHERE R.IDExpediente=" & p_IDExpediente & " "
            
    If UCase$(sTipo) = "MAIN" Then
        ' Adjudicatarios: raíz y contratista principal
        m_SQL = m_SQL & "AND (R.IdPadre Is Null AND R.ContratistaPrincipal='Sí') "
    Else
        ' Subcontratistas: hijos o raíz con SubContratista='Sí'
        m_SQL = m_SQL & "AND (R.IdPadre Is Not Null OR (R.IdPadre Is Null AND R.SubContratista='Sí')) "
    End If
    
    m_SQL = m_SQL & "ORDER BY T.Nemotecnico, T.Nombre;"
    
    ' Ejecución contra la BB.DD. de Expedientes
    Set rcd = getdb().OpenRecordset(m_SQL, dbOpenSnapshot)
    With rcd
        Do While Not .EOF
            ' Regla de visualización: primero Nemotécnico, si no hay usar Nombre
            sNombre = Nz(.fields("Nemotecnico"), "")
            If sNombre = "" Then
                sNombre = Nz(.fields("Nombre"), "")
            End If
            ' Limpieza por compatibilidad con controles (evitar ';')
            sNombre = Replace(sNombre, ";", ":")
            
            If sResultado = "" Then
                sResultado = sNombre
            Else
                sResultado = sResultado & "|" & sNombre
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rcd = Nothing
    GetCadenaJerarquicaEmpresas = Nz(sResultado, "")
    Exit Function
ManejoError:
    If Err.Number <> 0 Then
    p_Error = "Error en GetCadenaJerarquicaEmpresas (" & Err.Number & "): " & Err.description
    End If
End Function
'-------------------------------------------
' Nombre: EsContratistaPrincipal
' Propósito: Helper de validación para saber si un registro representa contratista principal
' Parámetros:
'   - vIdPadre As Variant: valor del campo IdPadre (Null para raíz)
'   - sContratistaPrincipal As String: valor del campo ContratistaPrincipal ('Sí'/'No')
' Retorno: Boolean (True si es raíz y ContratistaPrincipal='Sí')
'-------------------------------------------
Public Function EsContratistaPrincipal( _
                                        vIdPadre As Variant, _
                                        sContratistaPrincipal As String _
                                        ) As Boolean
    On Error GoTo ManejoError
    Dim esRaiz As Boolean
    Dim esPrincipal As Boolean
    
    esRaiz = (IsNull(vIdPadre))
    esPrincipal = (Trim$(Nz(sContratistaPrincipal, "")) = "Sí")
    
    EsContratistaPrincipal = (esRaiz And esPrincipal)
    Exit Function
ManejoError:
    ' En caso de error, devolver False
    EsContratistaPrincipal = False
End Function
'-------------------------------------------
' Nombre: EsSubContratista
' Propósito: Determinar si un registro representa un subcontratista según el modelo jerárquico
' Parámetros:
'   - vIdPadre As Variant: valor del campo IdPadre (Not Null para hijos; Null para raíz)
'   - sSubContratista As String: valor del campo SubContratista ('Sí'/'No')
' Retorno: Boolean
'   - True si es hijo (IdPadre Not Null) o raíz con SubContratista='Sí'
'   - False en caso contrario o si ocurre un error
'-------------------------------------------
Public Function EsSubContratista( _
                                vIdPadre As Variant, _
                                sSubContratista As String _
                                ) As Boolean
    On Error GoTo ManejoError
    Dim esHijo As Boolean
    Dim esRaizTelef As Boolean
    
    ' Es hijo si IdPadre NO es Null
    esHijo = (Not IsNull(vIdPadre))
    
    ' Es raíz Telefónica si IdPadre es Null y SubContratista='Sí'
    esRaizTelef = (IsNull(vIdPadre) And Trim$(Nz(sSubContratista, "")) = "Sí")
    
    ' Regla final del modelo jerárquico
    EsSubContratista = (esHijo Or esRaizTelef)
    Exit Function
ManejoError:
' En caso de error, devolver False
    EsSubContratista = False
End Function
'-------------------------------------------
' Nombre: getExpedienteSuministradores_RC
' Propósito: Devolver subcontratistas con retrocompatibilidad (bandera de entorno)
' Parámetros:
'   - p_IDExpediente As String
'   - p_Error As String (ByRef, opcional)
' Retorno: Scripting.Dictionary (IDSuministrador -> Suministrador)
'-------------------------------------------
Public Function getExpedienteSuministradores_RC( _
                                                p_IDExpediente As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
    On Error GoTo ManejoError
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Suministrador As Suministrador
    Dim sModelo As String
    
    If Nz(p_IDExpediente, "") = "" Then
        p_Error = "El parámetro p_IDExpediente es obligatorio."
        Exit Function
    End If
    
    sModelo = Trim$(Nz(Application.TempVars("CadenaJerarquicaModelo"), "nuevo"))
    
    If sModelo = "antiguo" Then
        ' Lógica previa (obsoleta): pierde hijos del árbol
        m_SQL = "SELECT T.* " & _
                "FROM TbSuministradores T INNER JOIN TbExpedientesSuministradores R " & _
                "ON T.IDSuministrador = R.IDSuministrador " & _
                "WHERE R.IDExpediente=" & p_IDExpediente & " " & _
                "AND R.SubContratista='Sí';"
    Else
        ' Lógica nueva (SUB): hijos o raíz con SubContratista='Sí'
        m_SQL = "SELECT T.* " & _
                "FROM TbExpedientesSuministradores R " & _
                "INNER JOIN TbSuministradores T ON R.IDSuministrador = T.IDSuministrador " & _
                "WHERE R.IDExpediente=" & p_IDExpediente & " " & _
                "AND (R.IdPadre Is Not Null OR (R.IdPadre Is Null AND R.SubContratista='Sí'));"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL, dbOpenSnapshot)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        
        .MoveFirst
        Do While Not .EOF
            Set m_Suministrador = New Suministrador
            For Each m_Campo In m_Suministrador.ColCampos
                m_Suministrador.SetPropiedad m_Campo, Nz(.fields(CStr(m_Campo)).value, ""), p_Error
                If Nz(p_Error, "") <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getExpedienteSuministradores_RC Is Nothing Then
                Set getExpedienteSuministradores_RC = New Scripting.Dictionary
                getExpedienteSuministradores_RC.CompareMode = TextCompare
            End If
            
            If Not getExpedienteSuministradores_RC.Exists(CStr(m_Suministrador.idSuministrador)) Then
                getExpedienteSuministradores_RC.Add CStr(m_Suministrador.idSuministrador), m_Suministrador
            End If
            
            .MoveNext
        Loop
    End With
    
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
ManejoError:
    If Err.Number <> 0 Then
        p_Error = "Error en getExpedienteSuministradores_RC (" & Err.Number & "): " & Err.description
    End If
End Function
'-------------------------------------------
' Nombre: getExpedienteJuridicas_RC
' Propósito: Devolver adjudicatarios (raíz) con retrocompatibilidad (bandera de entorno)
' Parámetros:
'   - p_IDExpediente As String
'   - p_Error As String (ByRef, opcional)
' Retorno: Scripting.Dictionary (IDSuministrador -> Suministrador)
'-------------------------------------------
Public Function getExpedienteJuridicas_RC( _
                                            p_IDExpediente As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    On Error GoTo ManejoError
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Suministrador As Suministrador
    Dim sModelo As String
    
    If Nz(p_IDExpediente, "") = "" Then
        p_Error = "El parámetro p_IDExpediente es obligatorio."
        Exit Function
    End If
    
    sModelo = Trim$(Nz(Application.TempVars("CadenaJerarquicaModelo"), "nuevo"))
    
    If sModelo = "antiguo" Then
        ' Lógica previa: adjudicatarios sin verificar raíz
        m_SQL = "SELECT T.* " & _
                "FROM TbExpedientesSuministradores R " & _
                "INNER JOIN TbSuministradores T ON R.IDSuministrador = T.IDSuministrador " & _
                "WHERE R.IDExpediente=" & p_IDExpediente & " " & _
                "AND R.ContratistaPrincipal='Sí';"
    Else
        ' Lógica nueva (MAIN): raíz + contratista principal
        m_SQL = "SELECT T.* " & _
                "FROM TbExpedientesSuministradores R " & _
                "INNER JOIN TbSuministradores T ON R.IDSuministrador = T.IDSuministrador " & _
                "WHERE R.IDExpediente=" & p_IDExpediente & " " & _
                "AND (R.IdPadre Is Null AND R.ContratistaPrincipal='Sí');"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL, dbOpenSnapshot)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        
        .MoveFirst
        Do While Not .EOF
            Set m_Suministrador = New Suministrador
            For Each m_Campo In m_Suministrador.ColCampos
                m_Suministrador.SetPropiedad m_Campo, Nz(.fields(CStr(m_Campo)).value, ""), p_Error
                If Nz(p_Error, "") <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getExpedienteJuridicas_RC Is Nothing Then
                Set getExpedienteJuridicas_RC = New Scripting.Dictionary
                getExpedienteJuridicas_RC.CompareMode = TextCompare
            End If
            
            If Not getExpedienteJuridicas_RC.Exists(CStr(m_Suministrador.idSuministrador)) Then
                getExpedienteJuridicas_RC.Add CStr(m_Suministrador.idSuministrador), m_Suministrador
            End If
            
            .MoveNext
        Loop
    End With
    
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
ManejoError:
    If Err.Number <> 0 Then
        p_Error = "Error en getExpedienteJuridicas_RC (" & Err.Number & "): " & Err.description
    End If
End Function
Public Function getRiesgosPorPrioridad( _
                                        p_IDEdicion As String, _
                                        Optional p_MostrandoRetirados As EnumSiNo = EnumSiNo.Sí, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    If p_IDEdicion = "" Then
        Exit Function
    End If
    If p_MostrandoRetirados <> EnumSiNo.Sí And p_MostrandoRetirados <> EnumSiNo.No Then
        p_MostrandoRetirados = EnumSiNo.Sí
    End If
    If p_MostrandoRetirados = EnumSiNo.Sí Then
        m_SQL = "SELECT * " & _
                "FROM TbRiesgos " & _
                "WHERE IDEdicion =" & p_IDEdicion & " " & _
                "ORDER BY Priorizacion;"
    Else
        m_SQL = "SELECT * " & _
                    "FROM TbRiesgos " & _
                    "WHERE FechaAprobacionRetiroPorCalidad Is Null " & _
                    "And IDEdicion =" & p_IDEdicion & " " & _
                    "ORDER BY Priorizacion;"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosPorPrioridad Is Nothing Then
                Set getRiesgosPorPrioridad = New Scripting.Dictionary
                getRiesgosPorPrioridad.CompareMode = TextCompare
            End If
            If Not getRiesgosPorPrioridad.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                getRiesgosPorPrioridad.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosPorPrioridad ha devuelto el error: " & Err.description
    End If
End Function

Public Function getRiesgosPorPrioridadTodos( _
                                                p_IDEdicion As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    If p_IDEdicion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDEdicion = " & p_IDEdicion & " And Not Priorizacion Is Null " & _
            "ORDER BY Priorizacion;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_ObjRiesgo = New riesgo
                For Each m_Campo In m_ObjRiesgo.ColCampos
                    m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getRiesgosPorPrioridadTodos Is Nothing Then
                    Set getRiesgosPorPrioridadTodos = New Scripting.Dictionary
                    getRiesgosPorPrioridadTodos.CompareMode = TextCompare
                End If
                If Not getRiesgosPorPrioridadTodos.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                    getRiesgosPorPrioridadTodos.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
                End If
                Set m_ObjRiesgo = Nothing
                .MoveNext
            Loop
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDEdicion= " & p_IDEdicion & " And Priorizacion Is Null " & _
            "ORDER BY CodigoRiesgo;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_ObjRiesgo = New riesgo
                For Each m_Campo In m_ObjRiesgo.ColCampos
                    m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getRiesgosPorPrioridadTodos Is Nothing Then
                    Set getRiesgosPorPrioridadTodos = New Scripting.Dictionary
                    getRiesgosPorPrioridadTodos.CompareMode = TextCompare
                End If
                If Not getRiesgosPorPrioridadTodos.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                    getRiesgosPorPrioridadTodos.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
                End If
                Set m_ObjRiesgo = Nothing
                .MoveNext
            Loop
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosPorPrioridadTodos ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgosPorProyecto( _
                                        p_IDProyecto As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_CodUnico As String
    Dim m_ObjRiesgo As riesgo
   
    On Error GoTo errores
    
    If p_IDProyecto = "" Then
        p_Error = "Se ha de indicar el IDProyecto"
        Err.Raise 1000
    End If
    m_SQL = "SELECT distinct TbRiesgos.CodigoUnico " & _
            "FROM TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
            "WHERE (((TbProyectosEdiciones.IDProyecto)=" & p_IDProyecto & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_CodUnico = .fields("CodigoUnico")
            Set m_ObjRiesgo = Constructor.getPrimerRiesgoByCodUnico(m_CodUnico, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If getRiesgosPorProyecto Is Nothing Then
                Set getRiesgosPorProyecto = New Scripting.Dictionary
                getRiesgosPorProyecto.CompareMode = TextCompare
            End If
            If Not getRiesgosPorProyecto.Exists(m_ObjRiesgo.IDRiesgo) Then
                getRiesgosPorProyecto.Add m_ObjRiesgo.IDRiesgo, m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgosPorProyecto ha devuelto el error: " & Err.description
    End If
End Function



Public Function getRiesgosTodosPorPrioridad( _
                                                p_IDEdicion As String, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    If p_IDEdicion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
                    "FROM TbRiesgos " & _
                    "WHERE IDEdicion=" & p_IDEdicion & " " & _
                    "ORDER BY Priorizacion;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosTodosPorPrioridad Is Nothing Then
                Set getRiesgosTodosPorPrioridad = New Scripting.Dictionary
                getRiesgosTodosPorPrioridad.CompareMode = TextCompare
            End If
            If Not getRiesgosTodosPorPrioridad.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                getRiesgosTodosPorPrioridad.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgosTodosPorPrioridad ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgosTodosNoOrdenados( _
                                                    p_IDEdicion As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    If p_IDEdicion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
                    "FROM TbRiesgos " & _
                    "WHERE IDEdicion=" & p_IDEdicion & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosTodosNoOrdenados Is Nothing Then
                Set getRiesgosTodosNoOrdenados = New Scripting.Dictionary
                getRiesgosTodosNoOrdenados.CompareMode = TextCompare
            End If
            If Not getRiesgosTodosNoOrdenados.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                getRiesgosTodosNoOrdenados.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgosTodosNoOrdenados ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgosNoRetiradosPorPrioridad( _
                                                    p_IDEdicion As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    If p_IDEdicion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
                    "FROM TbRiesgos " & _
                    "WHERE Not Mitigacion='Aceptar' " & _
                    "AND FechaRetirado Is Null " & _
                    "AND IDEdicion=" & p_IDEdicion & " " & _
                    "ORDER BY Priorizacion;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosNoRetiradosPorPrioridad Is Nothing Then
                Set getRiesgosNoRetiradosPorPrioridad = New Scripting.Dictionary
                getRiesgosNoRetiradosPorPrioridad.CompareMode = TextCompare
            End If
            If Not getRiesgosNoRetiradosPorPrioridad.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                getRiesgosNoRetiradosPorPrioridad.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgosNoRetiradosPorPrioridad ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgosNoRetiradosNoOrdenados( _
                                                    p_IDEdicion As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    If p_IDEdicion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
                    "FROM TbRiesgos " & _
                    "WHERE FechaRetirado Is Null " & _
                    "AND IDEdicion=" & p_IDEdicion & " " & _
                    "ORDER BY Priorizacion;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosNoRetiradosNoOrdenados Is Nothing Then
                Set getRiesgosNoRetiradosNoOrdenados = New Scripting.Dictionary
                getRiesgosNoRetiradosNoOrdenados.CompareMode = TextCompare
            End If
            If Not getRiesgosNoRetiradosNoOrdenados.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                getRiesgosNoRetiradosNoOrdenados.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgosNoRetiradosNoOrdenados ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgosMaterializados( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    
    m_SQL = "SELECT TbRiesgos.* " & _
            "FROM TbProyectos INNER JOIN (TbProyectosEdiciones INNER JOIN TbRiesgos " & _
            "ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion) " & _
            "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
            "WHERE ((Not (TbRiesgos.FechaMaterializado) Is Null)  " & _
            "AND ((TbProyectos.FechaCierre) Is Null) AND ((TbProyectosEdiciones.FechaPublicacion) Is Null));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosMaterializados Is Nothing Then
                Set getRiesgosMaterializados = New Scripting.Dictionary
                getRiesgosMaterializados.CompareMode = TextCompare
            End If
            If Not getRiesgosMaterializados.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                getRiesgosMaterializados.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgosMaterializados ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgosPorEdicionActivos( _
                                                p_IDEdicion As String, _
                                                Optional p_PorPriorizacion As EnumSiNo, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary

    Dim m_ColRiesgosTotales As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Riesgo As riesgo
    Dim m_Estado As EnumRiesgoEstado
    
    
    On Error GoTo errores
    If p_IDEdicion = "" Then
        Exit Function
    End If
    Set m_ColRiesgosTotales = Constructor.getRiesgosPorEdicion(p_IDEdicion, p_PorPriorizacion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColRiesgosTotales Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ColRiesgosTotales
        Set m_Riesgo = m_ColRiesgosTotales(m_ID)
        m_Riesgo.GrabarEstadoCalculado
        m_Estado = m_Riesgo.EstadoEnum
        If Not m_Estado = EnumRiesgoEstado.Retirado Then
            If getRiesgosPorEdicionActivos Is Nothing Then
                Set getRiesgosPorEdicionActivos = New Scripting.Dictionary
                getRiesgosPorEdicionActivos.CompareMode = TextCompare
            End If
            If Not getRiesgosPorEdicionActivos.Exists(CStr(m_Riesgo.IDRiesgo)) Then
                getRiesgosPorEdicionActivos.Add CStr(m_Riesgo.IDRiesgo), m_Riesgo
            End If
        End If
        
        
        Set m_Riesgo = Nothing
    Next
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgosPorEdicionActivos ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgosPorEdicion( _
                                        p_IDEdicion As String, _
                                        Optional p_PorPriorizacion As EnumSiNo, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_OrderBY As String
    
    On Error GoTo errores
    If p_IDEdicion = "" Then
        Exit Function
    End If
    If p_PorPriorizacion = EnumSiNo.Sí Then
       m_OrderBY = "ORDER BY Priorizacion;"
    Else
        m_OrderBY = "ORDER BY IDRiesgo;"
    End If
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDEdicion=" & p_IDEdicion & " " & _
            m_OrderBY
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjRiesgo = New riesgo
            For Each m_Campo In m_ObjRiesgo.ColCampos
                m_ObjRiesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getRiesgosPorEdicion Is Nothing Then
                Set getRiesgosPorEdicion = New Scripting.Dictionary
                getRiesgosPorEdicion.CompareMode = TextCompare
            End If
            If Not getRiesgosPorEdicion.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                getRiesgosPorEdicion.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
            End If
            Set m_ObjRiesgo = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgosPorEdicion ha devuelto el error: " & Err.description
    End If
End Function
Public Function getProyectosAbiertosParaIndicador( _
                                                    Optional p_Año As String, _
                                                    Optional ByVal p_Semestre As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
   
    
    Dim m_FechaInicial As String
    Dim m_FechaFinal As String
    Dim m_objColProyectos As Scripting.Dictionary
    Dim m_ObjProyecto As Proyecto
    Dim m_ID As Variant
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    
    p_Error = ""
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    If Not IsNumeric(p_Año) Then
        p_Año = Year(m_FechaRef)
    End If
    If p_Semestre <> "" Then
        If p_Semestre <> "1" And p_Semestre <> "2" Then
            p_Error = "Sólo hay dos semestres en un año"
            Err.Raise 1000
        End If
        If p_Semestre = "1" Then
            m_FechaInicial = "01/01/" & Format(p_Año, "0000")
            m_FechaFinal = "30/06/" & Format(p_Año, "0000")
        Else
            m_FechaInicial = "01/07/" & Format(p_Año, "0000")
            m_FechaFinal = Format("31/12/" & p_Año)
        End If
    Else
        m_FechaInicial = "01/01/" & Format(p_Año, "0000")
        m_FechaFinal = Format("31/12/" & p_Año)
    End If
    
    Set m_objColProyectos = Constructor.getProyectos(p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objColProyectos Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_objColProyectos
        Set m_ObjProyecto = m_objColProyectos(m_ID)
        If m_ObjProyecto.ParaInformeAvisos <> "Sí" Then
            GoTo siguiente
        End If
        If EstaEnElIntervaloDado(m_FechaInicial, m_FechaFinal, m_ObjProyecto.fechaRegistroInicial, m_ObjProyecto.FechaCierre) = EnumSiNo.Sí Then
            If getProyectosAbiertosParaIndicador Is Nothing Then
                Set getProyectosAbiertosParaIndicador = New Scripting.Dictionary
                getProyectosAbiertosParaIndicador.CompareMode = TextCompare
            End If
            If Not getProyectosAbiertosParaIndicador.Exists(m_ID) Then
                getProyectosAbiertosParaIndicador.Add m_ID, m_ObjProyecto
            End If
        End If
siguiente:
        Set m_ObjProyecto = Nothing
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getProyectosAbiertosParaIndicador ha devuelto el error: " & Err.description
    End If
End Function


Public Function getUltimoProyecto( _
                                    Optional p_IDUltimoProyecto As String, _
                                    Optional p_UsuarioRed As String, _
                                    Optional ByRef p_Error As String _
                                    ) As UltimoProyecto
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDUltimoProyecto = "" And p_UsuarioRed = "" Then
        Exit Function
    End If
    If p_IDUltimoProyecto <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbUltimoProyecto " & _
                "WHERE IDUltimoProyecto=" & p_IDUltimoProyecto & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbUltimoProyecto " & _
                "WHERE Usuario='" & p_UsuarioRed & "';"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getUltimoProyecto = New UltimoProyecto
        For Each m_Campo In getUltimoProyecto.ColCampos
            getUltimoProyecto.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getUltimoProyecto ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function



Public Function getCambiosPorEdicion( _
                                    p_IDProyecto As String, _
                                    p_Edicion As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Cambio As Cambio
    Dim i As Integer
    
    On Error GoTo errores
    
    If p_IDProyecto = "" And p_Edicion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM tbCambios " & _
            "WHERE IDProyecto=" & p_IDProyecto & " AND EdicionFinal=" & p_Edicion & " ORDER BY EdicionFinal;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Cambio = New Cambio
'            VBA.DoEvents
'            Debug.Print m_Cambio.IDCambio
'            VBA.DoEvents
            For Each m_Campo In m_Cambio.ColCampos
                m_Cambio.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getCambiosPorEdicion Is Nothing Then
                Set getCambiosPorEdicion = New Scripting.Dictionary
                getCambiosPorEdicion.CompareMode = TextCompare
            End If
            If Not getCambiosPorEdicion.Exists(CStr(m_Cambio.IDCambio)) Then
                getCambiosPorEdicion.Add CStr(m_Cambio.IDCambio), m_Cambio
            End If
            Set m_Cambio = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
            
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getCambiosPorEdicion ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getCambiosPorProyecto( _
                                        p_IDProyecto As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Cambio As Cambio
    Dim i As Integer
    
    On Error GoTo errores
    
    If p_IDProyecto = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM tbCambios " & _
            "WHERE IDProyecto=" & p_IDProyecto & " ORDER BY EdicionFinal;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Cambio = New Cambio
'            VBA.DoEvents
'            Debug.Print m_Cambio.IDCambio
'            VBA.DoEvents
            For Each m_Campo In m_Cambio.ColCampos
                m_Cambio.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getCambiosPorProyecto Is Nothing Then
                Set getCambiosPorProyecto = New Scripting.Dictionary
                getCambiosPorProyecto.CompareMode = TextCompare
            End If
            If Not getCambiosPorProyecto.Exists(CStr(m_Cambio.IDCambio)) Then
                getCambiosPorProyecto.Add CStr(m_Cambio.IDCambio), m_Cambio
            End If
            Set m_Cambio = Nothing
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
            
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getCambiosPorProyecto ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getNCs( _
                        Optional ByRef p_Error As String _
                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_NC As NC
    On Error GoTo errores
    
    
    m_SQL = "SELECT * " & _
            "FROM TbNoConformidades " & _
            "order by IDNoConformidad DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_NC = New NC
            For Each m_Campo In m_NC.ColCampos
                m_NC.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getNCs Is Nothing Then
                Set getNCs = New Scripting.Dictionary
                getNCs.CompareMode = TextCompare
            End If
            If Not getNCs.Exists(m_NC.IDNoConformidad) Then
                getNCs.Add m_NC.IDNoConformidad, m_NC
            End If
            Set m_NC = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getNCs ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getTiposNC( _
                        Optional ByRef p_Error As String _
                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    
    m_SQL = "SELECT * " & _
            "FROM TbTipologia;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            
            If getTiposNC Is Nothing Then
                Set getTiposNC = New Scripting.Dictionary
                getTiposNC.CompareMode = TextCompare
            End If
            If Not getTiposNC.Exists(.fields("CodTipologia").value) Then
                getTiposNC.Add .fields("CodTipologia").value, .fields("Tipologia").value
            End If
           
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getTiposNC ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getCarenciasExplicacion( _
                                        Optional p_Id As String, _
                                        Optional p_Apartado As String, _
                                        Optional ByRef p_Error As String _
                                        ) As CarenciasExplicacion
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    If p_Id = "" And p_Apartado = "" Then
        Exit Function
    End If
    If p_Id <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbExplicacionCarencias " & _
                "WHERE ID=" & p_Id & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbExplicacionCarencias " & _
                "WHERE Apartado='" & p_Apartado & "';"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getCarenciasExplicacion = New CarenciasExplicacion
        For Each m_Campo In getCarenciasExplicacion.ColCampos
            getCarenciasExplicacion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getCarenciasExplicacion ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getCarenciasExplicaciones( _
                                            Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Carencia As CarenciasExplicacion
   
    
    
    On Error GoTo errores
    
    m_SQL = "TbExplicacionCarencias"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Carencia = New CarenciasExplicacion
            For Each m_Campo In m_Carencia.ColCampos
                m_Carencia.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getCarenciasExplicaciones Is Nothing Then
                Set getCarenciasExplicaciones = New Scripting.Dictionary
                getCarenciasExplicaciones.CompareMode = TextCompare
            End If
            If Not getCarenciasExplicaciones.Exists(m_Carencia.ID) Then
                getCarenciasExplicaciones.Add m_Carencia.ID, m_Carencia
            End If
            Set m_Carencia = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getCarenciasExplicaciones ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getVersionesAplicacion( _
                                            Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Version As CCVersion
   
    
    
    On Error GoTo errores
    
    m_SQL = "SELECT * " & _
            "FROM TbAplicacionesVersiones " & _
            "WHERE IDAplicacion=" & m_IDAplicacion & " " & _
            "ORDER BY IDVersion DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Version = New CCVersion
            For Each m_Campo In m_Version.ColCampos
                m_Version.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getVersionesAplicacion Is Nothing Then
                Set getVersionesAplicacion = New Scripting.Dictionary
                getVersionesAplicacion.CompareMode = TextCompare
            End If
            If Not getVersionesAplicacion.Exists(m_Version.IDVersion) Then
                getVersionesAplicacion.Add m_Version.IDVersion, m_Version
            End If
            Set m_Version = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getVersionesAplicacion ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getCambiosExplicaciones( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_CambioExplicacion As CambioExplicacion
    
    On Error GoTo errores
    
    m_SQL = "TbCambiosExplicacion"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_CambioExplicacion = New CambioExplicacion
            For Each m_Campo In m_CambioExplicacion.ColCampos
                m_CambioExplicacion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getCambiosExplicaciones Is Nothing Then
                Set getCambiosExplicaciones = New Scripting.Dictionary
                getCambiosExplicaciones.CompareMode = TextCompare
            End If
            If Not getCambiosExplicaciones.Exists(m_CambioExplicacion.Apartado) Then
                getCambiosExplicaciones.Add m_CambioExplicacion.Apartado, m_CambioExplicacion
            End If
            Set m_CambioExplicacion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getCambiosExplicaciones ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function getEdicionesCaducadas( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Edicion As Edicion
        
    On Error GoTo errores
    
     m_SQL = "SELECT TbProyectosEdiciones.* " & _
            "FROM TbProyectos INNER JOIN TbProyectosEdiciones " & _
            "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
            "WHERE (((TbProyectos.FechaCierre) Is Null) " & _
            "AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
            "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null) " & _
            "AND ((DateDiff('d',Now(),[TbProyectosEdiciones].[FechaMaxProximaPublicacion]))<0));"
    
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
                m_Edicion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getEdicionesCaducadas Is Nothing Then
                Set getEdicionesCaducadas = New Scripting.Dictionary
                getEdicionesCaducadas.CompareMode = TextCompare
            End If
            If Not getEdicionesCaducadas.Exists(CStr(m_Edicion.IDEdicion)) Then
                getEdicionesCaducadas.Add CStr(m_Edicion.IDEdicion), m_Edicion
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
        p_Error = "El método constructor.getEdicionesCaducadas ha devuelto el error: " & Err.description
    End If
End Function
Public Function getEdicionesApuntoDeCaducar( _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Edicion As Edicion
        
    On Error GoTo errores
    
     m_SQL = "SELECT TbProyectosEdiciones.* " & _
            "FROM TbProyectos INNER JOIN TbProyectosEdiciones ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
            "WHERE (((TbProyectos.FechaCierre) Is Null) AND ((TbProyectosEdiciones.FechaPublicacion) Is Null) " & _
            "AND ((TbProyectosEdiciones.PropuestaRechazadaPorCalidadFecha) Is Null) " & _
            "AND ((DateDiff('d',Now(),[TbProyectos].[FechaMaxProximaPublicacion])) Between 1 And " & _
            m_ObjEntorno.JPDiasPreviosParaElAviso & "));"
    
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
                m_Edicion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            
            If getEdicionesApuntoDeCaducar Is Nothing Then
                Set getEdicionesApuntoDeCaducar = New Scripting.Dictionary
                getEdicionesApuntoDeCaducar.CompareMode = TextCompare
            End If
            If Not getEdicionesApuntoDeCaducar.Exists(CStr(m_Edicion.IDEdicion)) Then
                getEdicionesApuntoDeCaducar.Add CStr(m_Edicion.IDEdicion), m_Edicion
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
        p_Error = "El método constructor.getEdicionesApuntoDeCaducar ha devuelto el error: " & Err.description
    End If
End Function

Public Function getExpedientesBusqueda( _
                                        Optional p_PalabraClave As String, _
                                        Optional p_IDResponsableCalidad As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Expediente As Expediente
    
    On Error GoTo errores
    If p_IDResponsableCalidad <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbExpedientes " & _
                "WHERE IDResponsableCalidad =" & p_IDResponsableCalidad & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbExpedientes;"
    End If
       
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_Expediente = New Expediente
                For Each m_Campo In m_Expediente.ColCampos
                    m_Expediente.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If p_PalabraClave <> "" Then
                    If p_PalabraClave <> m_Expediente.IDExpediente And _
                        InStr(1, m_Expediente.Nemotecnico, p_PalabraClave) = 0 And _
                        InStr(1, m_Expediente.Titulo, p_PalabraClave) = 0 Then
                        GoTo siguiente
                    End If
                End If

                If getExpedientesBusqueda Is Nothing Then
                    Set getExpedientesBusqueda = New Scripting.Dictionary
                    getExpedientesBusqueda.CompareMode = TextCompare
                End If
                If Not getExpedientesBusqueda.Exists(CStr(m_Expediente.IDExpediente)) Then
                    getExpedientesBusqueda.Add CStr(m_Expediente.IDExpediente), m_Expediente
                End If
siguiente:
                .MoveNext
            Loop
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getExpedientesBusqueda ha devuelto el error: " & Err.description
    End If
End Function



Public Function getRiesgosBibliotecasFamilias( _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Familia As String
    
    On Error GoTo errores
    
    
    m_SQL = "SELECT distinct Familia " & _
            "FROM TbBibliotecaRiesgos " & _
            "WHERE Not Familia Is Null;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_Familia = Nz(.fields("Familia"), "")
            If m_Familia <> "" Then
                If getRiesgosBibliotecasFamilias Is Nothing Then
                    Set getRiesgosBibliotecasFamilias = New Scripting.Dictionary
                    getRiesgosBibliotecasFamilias.CompareMode = TextCompare
                End If
                If Not getRiesgosBibliotecasFamilias.Exists(m_Familia) Then
                    getRiesgosBibliotecasFamilias.Add m_Familia, m_Familia
                End If
            End If
            
            
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgosBibliotecasFamilias ha devuelto el error: " & Err.description
    End If
End Function
Public Function getSuministrador( _
                                    Optional p_Id As String, _
                                    Optional p_Nombre As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Suministrador
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_Id = "" And p_Nombre = "" Then
        Exit Function
    End If
    If p_Id <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbSuministradores " & _
                "WHERE IDSuministrador=" & p_Id & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbSuministradores " & _
                "WHERE Nombre='" & p_Nombre & "';"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getSuministrador = New Suministrador
        For Each m_Campo In getSuministrador.ColCampos
            getSuministrador.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getSuministrador ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function




Public Function getHistorialMaterializacionesRiesgo( _
                                                    p_Riesgo As riesgo, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Riesgo As riesgo
    Dim m_ColIDRiesgoFecha As Scripting.Dictionary
    On Error GoTo errores
    
    
    m_SQL = "SELECT TbRiesgos.* " & _
            "FROM TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
            "WHERE (((TbProyectosEdiciones.IDProyecto)=" & p_Riesgo.Edicion.IDProyecto & ") " & _
            "AND ((TbRiesgos.CodigoUnico)='" & p_Riesgo.CodigoUnico & "') " & _
            "AND (Not (TbRiesgos.FechaMaterializado) Is Null) " & _
            "AND ((TbProyectosEdiciones.IDEdicion)<" & p_Riesgo.IDEdicion & "));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_Riesgo = New riesgo
                For Each m_Campo In m_Riesgo.ColCampos
                    m_Riesgo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If m_ColIDRiesgoFecha Is Nothing Then
                    Set m_ColIDRiesgoFecha = New Scripting.Dictionary
                End If
                m_ColIDRiesgoFecha.CompareMode = TextCompare
                If Not m_ColIDRiesgoFecha.Exists(m_Riesgo.CodigoRiesgo & "|" & m_Riesgo.FechaMaterializado) Then
                    m_ColIDRiesgoFecha.Add m_Riesgo.CodigoRiesgo & "|" & m_Riesgo.FechaMaterializado, m_Riesgo.IDRiesgo & "|" & m_Riesgo.FechaMaterializado
                    If getHistorialMaterializacionesRiesgo Is Nothing Then
                        Set getHistorialMaterializacionesRiesgo = New Scripting.Dictionary
                        getHistorialMaterializacionesRiesgo.CompareMode = TextCompare
                    End If
                    If Not getHistorialMaterializacionesRiesgo.Exists(m_Riesgo.IDRiesgo) Then
                        getHistorialMaterializacionesRiesgo.Add m_Riesgo.IDRiesgo, m_Riesgo
                    End If
                End If
                
                Set m_Riesgo = Nothing
                .MoveNext
            Loop
            
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getHistorialMaterializacionesRiesgo ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgoMaterializado( _
                                        p_Id As String, _
                                        Optional ByRef p_Error As String _
                                        ) As RiesgoMaterializacion
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If p_Id = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
                "FROM TbRiesgosMaterializaciones " & _
                "WHERE ID=" & p_Id & "; "
   
   Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getRiesgoMaterializado = New RiesgoMaterializacion
        For Each m_Campo In getRiesgoMaterializado.ColCampos
            getRiesgoMaterializado.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método getRiesgoMaterializado ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgoMaterializadoUltimo( _
                                                Optional p_IDProyecto As String, _
                                                Optional p_CodigoRiesgo As String, _
                                                Optional p_Riesgo As riesgo, _
                                                Optional ByRef p_Error As String _
                                                ) As RiesgoMaterializacion
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If p_Riesgo Is Nothing And (p_IDProyecto = "" Or p_CodigoRiesgo = "") Then
        Exit Function
    End If
    If p_IDProyecto = "" Or p_CodigoRiesgo = "" Then
        p_IDProyecto = p_Riesgo.Edicion.IDProyecto
        p_CodigoRiesgo = p_Riesgo.CodigoRiesgo
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE IDProyecto=" & p_IDProyecto & " " & _
            "AND CodigoRiesgo='" & p_CodigoRiesgo & "' " & _
            "ORDER BY ID DESC;"
   
   Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getRiesgoMaterializadoUltimo = New RiesgoMaterializacion
        For Each m_Campo In getRiesgoMaterializadoUltimo.ColCampos
            getRiesgoMaterializadoUltimo.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método getRiesgoMaterializadoUltimo ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgoMaterializadoEstadoAnterior( _
                                                        Optional p_RiesgoMaterializado As RiesgoMaterializacion, _
                                                        Optional ByRef p_Error As String _
                                                        ) As RiesgoMaterializacion
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_RiesgoMaterializadoActual As RiesgoMaterializacion
    
    
    On Error GoTo errores
    If p_RiesgoMaterializado Is Nothing Then
        Exit Function
    End If
    
    If Not IsNumeric(p_RiesgoMaterializado.ID) Then
        With p_RiesgoMaterializado
             Set getRiesgoMaterializadoEstadoAnterior = Constructor.getRiesgoMaterializadoUltimo( _
                                                    .IDProyecto, .CodigoRiesgo, .riesgo, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
                                    
        End With
       
        Exit Function
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE ID< " & p_RiesgoMaterializado.ID & " " & _
            "AND IDProyecto=" & p_RiesgoMaterializado.IDProyecto & " " & _
            "AND CodigoRiesgo='" & p_RiesgoMaterializado.CodigoRiesgo & "' " & _
            "ORDER BY ID DESC;"
   
   Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Set getRiesgoMaterializadoEstadoAnterior = New RiesgoMaterializacion
        For Each m_Campo In getRiesgoMaterializadoEstadoAnterior.ColCampos
            getRiesgoMaterializadoEstadoAnterior.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método getRiesgoMaterializadoEstadoAnterior ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgoMaterializadoHistoria( _
                                                Optional p_IDProyecto As String, _
                                                Optional p_CodigoRiesgo As String, _
                                                Optional p_Riesgo As riesgo, _
                                                Optional ByRef p_Error As String _
                                                ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_RiesgoMat As RiesgoMaterializacion
    
    On Error GoTo errores
    If p_Riesgo Is Nothing And (p_IDProyecto = "" Or p_CodigoRiesgo = "") Then
        Exit Function
    End If
    If p_IDProyecto = "" Or p_CodigoRiesgo = "" Then
        p_IDProyecto = p_Riesgo.Edicion.IDProyecto
        p_CodigoRiesgo = p_Riesgo.CodigoRiesgo
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE IDProyecto=" & p_IDProyecto & " " & _
            "AND CodigoRiesgo='" & p_CodigoRiesgo & "' " & _
            "ORDER BY ID Desc;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_RiesgoMat = New RiesgoMaterializacion
                For Each m_Campo In m_RiesgoMat.ColCampos
                    m_RiesgoMat.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getRiesgoMaterializadoHistoria Is Nothing Then
                    Set getRiesgoMaterializadoHistoria = New Scripting.Dictionary
                    getRiesgoMaterializadoHistoria.CompareMode = TextCompare
                End If
                If Not getRiesgoMaterializadoHistoria.Exists(m_RiesgoMat.ID) Then
                    getRiesgoMaterializadoHistoria.Add m_RiesgoMat.ID, m_RiesgoMat
                End If
                
                Set m_RiesgoMat = Nothing
                .MoveNext
            Loop
            
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgoMaterializadoHistoria ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRiesgoMaterializadoPendientesDecidirNC( _
                                                            Optional p_IDProyecto As String, _
                                                            Optional p_CodigoRiesgo As String, _
                                                            Optional p_Riesgo As riesgo, _
                                                            Optional ByRef p_Error As String _
                                                            ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_RiesgoMat As RiesgoMaterializacion
    
    On Error GoTo errores
    If p_Riesgo Is Nothing And (p_IDProyecto = "" Or p_CodigoRiesgo = "") Then
        Exit Function
    End If
    If p_IDProyecto = "" Or p_CodigoRiesgo = "" Then
        p_IDProyecto = p_Riesgo.Edicion.IDProyecto
        p_CodigoRiesgo = p_Riesgo.CodigoRiesgo
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE IDProyecto=" & p_IDProyecto & " " & _
            "AND CodigoRiesgo='" & p_CodigoRiesgo & "' " & _
            "ORDER BY Fecha;"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_RiesgoMat = New RiesgoMaterializacion
                For Each m_Campo In m_RiesgoMat.ColCampos
                    m_RiesgoMat.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getRiesgoMaterializadoPendientesDecidirNC Is Nothing Then
                    Set getRiesgoMaterializadoPendientesDecidirNC = New Scripting.Dictionary
                    getRiesgoMaterializadoPendientesDecidirNC.CompareMode = TextCompare
                End If
                If Not getRiesgoMaterializadoPendientesDecidirNC.Exists(m_RiesgoMat.ID) Then
                    getRiesgoMaterializadoPendientesDecidirNC.Add m_RiesgoMat.ID, m_RiesgoMat
                End If

                Set m_RiesgoMat = Nothing
                .MoveNext
            Loop

        End If

    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT TbRiesgosMaterializaciones.* " & _
            "FROM TbRiesgosNC INNER JOIN TbRiesgosMaterializaciones " & _
            "ON TbRiesgosNC.ID = TbRiesgosMaterializaciones.IDNC " & _
            "WHERE (((TbRiesgosMaterializaciones.IDProyecto)=" & p_IDProyecto & ") " & _
            "AND ((TbRiesgosMaterializaciones.CodigoRiesgo)='" & p_CodigoRiesgo & "') " & _
            "AND ((TbRiesgosMaterializaciones.EsMaterializacion)='Sí') " & _
            "AND ((TbRiesgosNC.FechaDecison) Is Null));"

    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_RiesgoMat = New RiesgoMaterializacion
                For Each m_Campo In m_RiesgoMat.ColCampos
                    m_RiesgoMat.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Next
                If getRiesgoMaterializadoPendientesDecidirNC Is Nothing Then
                    Set getRiesgoMaterializadoPendientesDecidirNC = New Scripting.Dictionary
                    getRiesgoMaterializadoPendientesDecidirNC.CompareMode = TextCompare
                End If
                If Not getRiesgoMaterializadoPendientesDecidirNC.Exists(m_RiesgoMat.ID) Then
                    getRiesgoMaterializadoPendientesDecidirNC.Add m_RiesgoMat.ID, m_RiesgoMat
                End If
                
                Set m_RiesgoMat = Nothing
                .MoveNext
            Loop
            
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getRiesgoMaterializadoPendientesDecidirNC ha devuelto el error: " & Err.description
    End If
End Function
Public Function EsPriorizacionCorrecta( _
                                        p_IDEdicion As String, _
                                        p_Ordinal As String, _
                                        Optional p_IDRiesgo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As EnumSiNo
    
    Dim m_EsPriorizacionUsada As EnumSiNo
    Dim m_EsPriorizacionFueraDeRango As EnumSiNo
    
    On Error GoTo errores
    
    m_EsPriorizacionUsada = EsPriorizacionUsada(p_IDEdicion, p_Ordinal, p_IDRiesgo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EsPriorizacionUsada = EnumSiNo.Sí Then
        EsPriorizacionCorrecta = EnumSiNo.No
        Exit Function
    End If
    
    
    m_EsPriorizacionFueraDeRango = EsPriorizacionFueraDeRango(p_IDEdicion, p_Ordinal, p_IDRiesgo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EsPriorizacionFueraDeRango = EnumSiNo.Sí Then
        EsPriorizacionCorrecta = EnumSiNo.No
    Else
        EsPriorizacionCorrecta = EnumSiNo.Sí
    End If
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EsPriorizacionCorrecta ha devuelto el error: " & Err.description
    End If
End Function
Private Function EsPriorizacionUsada( _
                                    p_IDEdicion As String, _
                                    p_Ordinal As String, _
                                    Optional p_IDRiesgo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As EnumSiNo
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If p_IDEdicion = "" Or p_Ordinal = "" Then
        Exit Function
    End If
    If p_IDRiesgo = "" Then
        m_SQL = "SELECT IDEdicion " & _
                "FROM TbRiesgos " & _
                "WHERE IDEdicion=" & p_IDEdicion & _
                " AND Priorizacion=" & p_Ordinal & ";"
    Else
        m_SQL = "SELECT IDEdicion " & _
                "FROM TbRiesgos " & _
                "WHERE IDEdicion=" & p_IDEdicion & " " & _
                "AND IDRiesgo<>" & p_IDRiesgo & " " & _
                "AND Priorizacion=" & p_Ordinal & ";"
    End If
    
   
   Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            EsPriorizacionUsada = EnumSiNo.No
        Else
            EsPriorizacionUsada = EnumSiNo.Sí
        End If
       
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EsPriorizacionUsada ha devuelto el error: " & Err.description
    End If
End Function
Private Function EsPriorizacionFueraDeRango( _
                                            p_IDEdicion As String, _
                                            p_Ordinal As String, _
                                            Optional p_IDRiesgo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As EnumSiNo
    
    Dim m_Edicion As Edicion
    
    
    On Error GoTo errores
    If Not IsNumeric(p_Ordinal) Then
        p_Error = "El Ordinal ha de ser un número"
        Err.Raise 1000
    End If
    If CInt(p_Ordinal) = 0 Then
        EsPriorizacionFueraDeRango = EnumSiNo.Sí
        Exit Function
    End If
    Set m_Edicion = Constructor.getEdicion(p_IDEdicion:=p_IDEdicion, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Edicion Is Nothing Then
        p_Error = "No se ha podido determinar la edición"
        Err.Raise 1000
    End If
    If m_Edicion.colRiesgosNoRetirados Is Nothing Then
        
        EsPriorizacionFueraDeRango = EnumSiNo.Sí
        Exit Function
    End If
    
    If CInt(p_Ordinal) > CInt(m_Edicion.colRiesgosNoRetirados.Count) Then
        If IsNumeric(p_IDRiesgo) Then
            EsPriorizacionFueraDeRango = EnumSiNo.Sí
        Else
            If CInt(p_Ordinal) = CInt(m_Edicion.colRiesgosNoRetirados.Count + 1) Then
                EsPriorizacionFueraDeRango = EnumSiNo.No
            Else
                EsPriorizacionFueraDeRango = EnumSiNo.Sí
            End If
        End If
    Else
        EsPriorizacionFueraDeRango = EnumSiNo.No
    End If
    
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EsPriorizacionFueraDeRango ha devuelto el error: " & Err.description
    End If
End Function
Public Function getCCVersion( _
                                Optional p_IDVersion As String, _
                                Optional p_Version As String, _
                                Optional ByRef p_Error As String _
                                ) As CCVersion
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDVersion = "" And p_Version = "" Then
        Exit Function
    End If
    If p_IDVersion <> "" Then
        m_SQL = "SELECT * " & _
            "FROM TbAplicacionesVersiones " & _
            "WHERE IDVersion=" & p_IDVersion & ";"
    Else
        m_SQL = "SELECT * " & _
            "FROM TbAplicacionesVersiones " & _
            "WHERE Version='" & p_Version & "';"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getCCVersion = New CCVersion
        
        For Each m_Campo In getCCVersion.ColCampos
            getCCVersion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método getCCVersion ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getCCVersiones( _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Version As CCVersion
    
    On Error GoTo errores
    
    If m_IDAplicacion = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * " & _
                "FROM TbAplicacionesVersiones " & _
                "WHERE IDAplicacion=" & m_IDAplicacion & ";"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Version = New CCVersion
            For Each m_Campo In m_Version.ColCampos
                m_Version.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getCCVersiones Is Nothing Then
                Set getCCVersiones = New Scripting.Dictionary
                getCCVersiones.CompareMode = TextCompare
            End If
            If Not getCCVersiones.Exists(m_Version.IDVersion) Then
                getCCVersiones.Add m_Version.IDVersion, m_Version
            End If
            Set m_Version = Nothing
            .MoveNext
        Loop
       
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método getCCVersiones ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getCCCambio( _
                                Optional p_IDCambio As String, _
                                Optional ByRef p_Error As String _
                                ) As CCCambio
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDCambio = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbVersionCambios " & _
            "WHERE IDCambio=" & p_IDCambio & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getCCCambio = New CCCambio
        For Each m_Campo In getCCCambio.ColCampos
            getCCCambio.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método getCCCambio ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getCCCambios( _
                                Optional p_IDVersion As String, _
                                Optional p_Version As String, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Cambio As CCCambio
    
    On Error GoTo errores
    
    If p_IDVersion = "" And p_Version = "" Then
        Exit Function
    End If
    If p_IDVersion <> "" Then
        m_SQL = "SELECT * FROM TbVersionCambios " & _
                "WHERE IDVersion=" & p_IDVersion & ";"
    Else
        m_SQL = "SELECT TbVersionCambios.* " & _
                "FROM TbAplicacionesVersiones INNER JOIN TbVersionCambios " & _
                "ON TbAplicacionesVersiones.IDVersion = TbVersionCambios.IDVersion " & _
                "WHERE (((TbAplicacionesVersiones.Version)='" & p_Version & "'));"
    End If
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Cambio = New CCCambio
            For Each m_Campo In m_Cambio.ColCampos
                m_Cambio.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getCCCambios Is Nothing Then
                Set getCCCambios = New Scripting.Dictionary
                getCCCambios.CompareMode = TextCompare
            End If
            If Not getCCCambios.Exists(m_Cambio.IDCambio) Then
                getCCCambios.Add m_Cambio.IDCambio, m_Cambio
            End If
            Set m_Cambio = Nothing
            .MoveNext
        Loop
       
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método getCCCambios ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getDocumentosCCCambio( _
                                    Optional p_IDCambio As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_DocCambio As CCDocumentoCambio
    
    On Error GoTo errores
    
    If p_IDCambio = "" Then
        Exit Function
    End If
    m_SQL = "SELECT * FROM TbCambioDocumentos " & _
            "WHERE IDCambio=" & p_IDCambio & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_DocCambio = New CCDocumentoCambio
            For Each m_Campo In m_DocCambio.ColCampos
                m_DocCambio.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getDocumentosCCCambio Is Nothing Then
                Set getDocumentosCCCambio = New Scripting.Dictionary
                getDocumentosCCCambio.CompareMode = TextCompare
            End If
            If Not getDocumentosCCCambio.Exists(m_DocCambio.IDDoc) Then
                getDocumentosCCCambio.Add m_DocCambio.IDDoc, m_DocCambio
            End If
            Set m_DocCambio = Nothing
            .MoveNext
        Loop
       
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método getDocumentosCCCambio ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function



Public Function getOrganoContratacion( _
                                    Optional p_IDOrganoContratacion As String, _
                                    Optional p_OrganoContratacion As String, _
                                    Optional ByRef p_Error As String _
                                    ) As OrganoContratacion
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    
    
    On Error GoTo errores
    
    If p_IDOrganoContratacion = "" And p_OrganoContratacion = "" Then
        Exit Function
    End If
    If p_IDOrganoContratacion <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbOrganosContratacion " & _
                "WHERE IDOrganoContratacion=" & p_IDOrganoContratacion & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbOrganosContratacion " & _
                "WHERE OrganoContratacion='" & p_OrganoContratacion & "';"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getOrganoContratacion = New OrganoContratacion
        For Each m_Campo In getOrganoContratacion.ColCampos
            getOrganoContratacion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método getOrganoContratacion ha devuelto el error: " & Err.description
    End If
End Function
Public Function getRAC( _
                            Optional p_IDRAC As String, _
                            Optional p_RAC As String, _
                            Optional ByRef p_Error As String _
                            ) As RAC
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    
    
    On Error GoTo errores
    
    If p_IDRAC = "" And p_RAC = "" Then
        Exit Function
    End If
    If p_IDRAC <> "" Then
        m_SQL = "SELECT * " & _
                "FROM TbRACS " & _
                "WHERE IDRAC=" & p_IDRAC & ";"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbRACS " & _
                "WHERE RAC='" & p_RAC & "';"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getRAC = New RAC
        For Each m_Campo In getRAC.ColCampos
            getRAC.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "El método getRAC ha devuelto el error: " & Err.description
    End If
End Function

Public Function getExpedienteRACS( _
                                    p_IDExpediente As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_RAC As RAC
    
    
    On Error GoTo errores
    
    If p_IDExpediente = "" Then
       Exit Function
    End If
    m_SQL = "SELECT TbRACS.* " & _
                "FROM TbExpedientesRACS INNER JOIN TbRACS ON TbExpedientesRACS.IDRAC = TbRACS.IDRAC " & _
                "WHERE TbExpedientesRACS.IDExpediente=" & p_IDExpediente & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_RAC = New RAC
            For Each m_Campo In m_RAC.ColCampos
                m_RAC.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                 If p_Error <> "" Then
                     Err.Raise 1000
                 End If
             Next
             If getExpedienteRACS Is Nothing Then
                Set getExpedienteRACS = New Scripting.Dictionary
                getExpedienteRACS.CompareMode = TextCompare
             End If
             If Not getExpedienteRACS.Exists(CStr(m_RAC.IDRAC)) Then
                getExpedienteRACS.Add CStr(m_RAC.IDRAC), m_RAC
             End If
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getExpedienteRACS ha devuelto el error: " & Err.description
    End If
End Function

Public Function getOrdinalSiguiente( _
                                    p_IDExpediente As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_OrdinalMax As Integer
    
    On Error GoTo errores
     
    If p_IDExpediente = "" Then
        Exit Function
    End If
    
    m_SQL = "SELECT Ordinal " & _
            "FROM TbProyectos " & _
            "WHERE IDExpediente=" & p_IDExpediente & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                If IsNumeric(Nz(.fields("Ordinal"), "")) Then
                    If CInt(.fields("Ordinal")) > m_OrdinalMax Then
                        m_OrdinalMax = CInt(.fields("Ordinal"))
                    End If
                End If
                .MoveNext
            Loop
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    getOrdinalSiguiente = CStr(m_OrdinalMax + 1)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método constructor.getOrdinalSiguiente ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPmAccionReversa( _
                                    p_IDAccionMitigacion As String, _
                                    Optional ByRef p_Error As String _
                                    ) As PMAccionReversa
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDAccionMitigacion = "" Then
        p_Error = "Falta el p_IDAccionMitigacion"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosPlanMitigacionDetalleReversa " & _
            "WHERE IDAccionMitigacion=" & p_IDAccionMitigacion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getPmAccionReversa = New PMAccionReversa
        For Each m_Campo In getPmAccionReversa.ColCampos
            getPmAccionReversa.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getPmAccionReversa ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPcAccionReversa( _
                                    p_IDAccionContingencia As String, _
                                    Optional ByRef p_Error As String _
                                    ) As PCAccionReversa
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    
    On Error GoTo errores
    
    If p_IDAccionContingencia = "" Then
        p_Error = "Falta el p_IDAccionContingencia"
        Err.Raise 1000
    End If
    m_SQL = "SELECT * " & _
            "FROM TbRiesgosPlanContingenciaDetalleReversa " & _
            "WHERE IDAccionContingencia=" & p_IDAccionContingencia & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        Set getPcAccionReversa = New PCAccionReversa
        For Each m_Campo In getPcAccionReversa.ColCampos
            getPcAccionReversa.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
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
        p_Error = "EL método constructor.getPcAccionReversa ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getEdicionesActivas( _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_objEdicion As Edicion
    
    On Error GoTo errores
    
   m_SQL = "SELECT TbProyectosEdiciones.* " & _
            "FROM TbProyectos INNER JOIN TbProyectosEdiciones " & _
            "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto " & _
            "WHERE (((TbProyectos.FechaCierre) Is Null) AND ((TbProyectosEdiciones.FechaPublicacion) Is Null));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_objEdicion = New Edicion
            For Each m_Campo In m_objEdicion.ColCampos
                m_objEdicion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getEdicionesActivas Is Nothing Then
                Set getEdicionesActivas = New Scripting.Dictionary
                getEdicionesActivas.CompareMode = TextCompare
            End If
            If Not getEdicionesActivas.Exists(m_objEdicion.IDEdicion) Then
                getEdicionesActivas.Add m_objEdicion.IDEdicion, m_objEdicion
            End If
            Set m_objEdicion = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getEdicionesActivas ha devuelto el error: " & Err.description
    End If
End Function

Public Function getPMAccionesSinCerrar( _
                                        p_IDRiesgo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Accion As PMAccion
    
    On Error GoTo errores
    
    If p_IDRiesgo = "" Then
        Exit Function
    End If
    
    
    m_SQL = "SELECT TbRiesgosPlanMitigacionDetalle.* " & _
            "FROM TbRiesgosPlanMitigacionPpal INNER JOIN TbRiesgosPlanMitigacionDetalle " & _
            "ON TbRiesgosPlanMitigacionPpal.IDMitigacion = TbRiesgosPlanMitigacionDetalle.IDMitigacion " & _
            "WHERE IDRiesgo=" & p_IDRiesgo & " AND FechaFinReal Is Null;"

    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Accion = New PMAccion
            For Each m_Campo In m_Accion.ColCampos
                m_Accion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getPMAccionesSinCerrar Is Nothing Then
                Set getPMAccionesSinCerrar = New Scripting.Dictionary
                getPMAccionesSinCerrar.CompareMode = TextCompare
            End If
            If Not getPMAccionesSinCerrar.Exists(m_Accion.IDAccionMitigacion) Then
                getPMAccionesSinCerrar.Add m_Accion.IDAccionMitigacion, m_Accion
            End If
            Set m_Accion = Nothing
            .MoveNext
        Loop
       
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método getPMAccionesSinCerrar ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function getPCAccionesSinCerrar( _
                                        p_IDRiesgo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Accion As PCAccion
    
    On Error GoTo errores
    
    If p_IDRiesgo = "" Then
        Exit Function
    End If
    
    
    m_SQL = "SELECT TbRiesgosPlanContingenciaDetalle.* " & _
            "FROM TbRiesgosPlanContingenciaPpal INNER JOIN TbRiesgosPlanContingenciaDetalle " & _
            "ON TbRiesgosPlanContingenciaPpal.IDContingencia = TbRiesgosPlanContingenciaDetalle.IDContingencia " & _
            "WHERE IDRiesgo=" & p_IDRiesgo & " AND FechaFinReal Is Null;"

    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_Accion = New PCAccion
            For Each m_Campo In m_Accion.ColCampos
                m_Accion.SetPropiedad m_Campo, Nz(.fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getPCAccionesSinCerrar Is Nothing Then
                Set getPCAccionesSinCerrar = New Scripting.Dictionary
                getPCAccionesSinCerrar.CompareMode = TextCompare
            End If
            If Not getPCAccionesSinCerrar.Exists(m_Accion.IDAccionContingencia) Then
                getPCAccionesSinCerrar.Add m_Accion.IDAccionContingencia, m_Accion
            End If
            Set m_Accion = Nothing
            .MoveNext
        Loop
       
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método getPCAccionesSinCerrar ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function






