Attribute VB_Name = "Módulo1"
Option Compare Database
Option Explicit



Public Function RegularizarTiemposAceptadosRetirados(Optional ByRef p_Error As String) As String
    
    Dim m_ColRiesgosAceptadosRetirados As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Riesgo As riesgo
    Dim m_TiempoCalidad As String
    Dim m_SQL As String
    Dim i As Long
    
    On Error GoTo errores
    Set m_ColRiesgosAceptadosRetirados = getRiesgosAceptadosORetiradosParaRegularizar(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColRiesgosAceptadosRetirados Is Nothing Then
        Exit Function
    End If
    i = 1
    For Each m_ID In m_ColRiesgosAceptadosRetirados
        Set m_Riesgo = m_ColRiesgosAceptadosRetirados(m_ID)
        VBA.DoEvents
        If IsDate(m_Riesgo.FechaJustificacionAceptacionRiesgo) Then
            m_Riesgo.RegistrarDiasAceptacionCalidad p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            VBA.DoEvents
            Debug.Print "regularizar Tiemps Aceptados: " & m_Riesgo.CodigoUnico, m_Riesgo.DiasRespuestaCalidadAceptacion & " (" & i & " de " & m_ColRiesgosAceptadosRetirados.Count & ")"
            VBA.DoEvents
            
        End If
        If IsDate(m_Riesgo.FechaJustificacionRetiroRiesgo) Then
            m_Riesgo.RegistrarDiasAceptacionRetiro p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            VBA.DoEvents
            Debug.Print "regularizar Tiemps Retirados: " & m_Riesgo.CodigoUnico, m_Riesgo.DiasRespuestaCalidadRetiro & " (" & i & " de " & m_ColRiesgosAceptadosRetirados.Count & ")"
            VBA.DoEvents
            
        End If
        
        i = i + 1
        Set m_Riesgo = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método RegularizarTiemposAceptados ha devuelto el error: " & vbNewLine & Err.description
        
    End If
    Debug.Print p_Error
End Function


Private Function getRiesgosAceptadosORetiradosParaRegularizar( _
                                                                Optional ByRef p_Error As String _
                                                                ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
        
    On Error GoTo errores
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE (" & _
                "Not FechaJustificacionAceptacionRiesgo Is Null " & _
                "AND Not FechaAprobacionAceptacionPorCalidad Is Null " & _
                "AND DiasSinRespuestaCalidadAceptacion Is Null" & _
                    ") " & _
                "OR (" & _
                    "Not FechaJustificacionRetiroRiesgo Is Null " & _
                    "AND Not FechaAprobacionRetiroPorCalidad Is Null " & _
                    "AND DiasSinRespuestaCalidadRetiro Is Null" & _
                    ");"
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
            If getRiesgosAceptadosORetiradosParaRegularizar Is Nothing Then
                Set getRiesgosAceptadosORetiradosParaRegularizar = New Scripting.Dictionary
                getRiesgosAceptadosORetiradosParaRegularizar.CompareMode = TextCompare
            End If
            If Not getRiesgosAceptadosORetiradosParaRegularizar.Exists(m_ObjRiesgo.IDRiesgo) Then
                getRiesgosAceptadosORetiradosParaRegularizar.Add m_ObjRiesgo.IDRiesgo, m_ObjRiesgo
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
        p_Error = "El método constructor.getRiesgosAceptadosORetiradosParaRegularizar ha devuelto el error: " & Err.description
    End If
End Function

Public Function RegularizarRiesgosMaterializados(Optional ByRef p_Error As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    Dim m_IdRiesgo As String
    Dim m_IDNC As String
    Dim m_ParaNC As String
    Dim m_FechaDecision As String
    
    
    Dim m_CodigoRiesgo As String
    Dim m_IDProyecto As String
    Dim m_IDEdicion As String
    
    
    Dim m_Riesgo As riesgo
    Dim m_RiesgoMaterializado As RiesgoMaterializacion
    
    Dim m_FechaMaterializado As String
    Dim m_Col As Scripting.Dictionary
    Dim m_ColRiesgosMaterializados As Scripting.Dictionary
    Dim m_ID As Variant
    Dim dato As Variant
    Dim i As Long
    Dim db As DAO.Database
    Dim wks As DAO.Workspace
    Dim blnTransStarted As Boolean
    
    On Error GoTo errores
    Set m_Col = getcolRiesgoNC(p_Error)
    If m_Col Is Nothing Then
        Exit Function
    End If
    m_SQL = "DELETE * " & _
            "FROM TbRiesgosMaterializaciones;"
    Set db = getdb()
    Set wks = DBEngine.Workspaces(0)
    wks.BeginTrans
    blnTransStarted = True
    db.Execute m_SQL, dbFailOnError
    i = 1
    For Each m_ID In m_Col
        dato = Split(m_ID, "|")
        m_IdRiesgo = dato(0)
        m_IDNC = dato(1)
        m_ParaNC = dato(2)
        m_FechaDecision = dato(3)
        Set m_Riesgo = Constructor.getRiesgo(m_IdRiesgo, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Riesgo Is Nothing Then
            GoTo siguiente
        End If
        If Not IsDate(m_Riesgo.FechaMaterializado) Then
            GoTo siguiente
        End If
        Set m_RiesgoMaterializado = New RiesgoMaterializacion
        With m_RiesgoMaterializado
            .IDProyecto = m_Riesgo.Edicion.IDProyecto
            .IDEdicion = m_Riesgo.IDEdicion
            .CodigoRiesgo = m_Riesgo.CodigoRiesgo
            .Fecha = m_Riesgo.FechaMaterializado
            .EsMaterializacion = "Sí"
            RegistrarMaterializacion m_RiesgoMaterializado, p_Error
            
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If IsNumeric(m_IDNC) Then
                .RegistrarParaNC m_IDNC, m_FechaDecision, p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Else
                If m_ParaNC = "No" Then
                    .RegistrarParaNONC m_FechaDecision, p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                End If
            End If
        End With
        VBA.DoEvents
        Debug.Print "RegularizarRiesgosMaterializados: " & m_Riesgo.CodigoUnico & " (" & i & " de " & m_Col.Count & ")"
        VBA.DoEvents
        'If m_Riesgo.CodigoUnico = "029R018" Then Stop
        i = i + 1
siguiente:
    Next
    wks.CommitTrans
    blnTransStarted = False
    
    Exit Function
errores:
    If blnTransStarted Then
        On Error Resume Next
        wks.Rollback
        On Error GoTo 0
    End If
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.RegularizarRiesgosMaterializados ha devuelto el error: " & Err.description
    End If
    
End Function

Private Function getcolRiesgosParaCambioCodigoCompleto(Optional ByRef p_Error As String) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_IdRiesgo As String
    Dim m_CodigoUnico As String
    
    On Error GoTo errores
    m_SQL = "SELECT TbRiesgos.IDRiesgo, Format([IDProyecto],'000') & [CodigoRiesgo] AS CodUnico " & _
            "FROM TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
            "WHERE (((Format([IDProyecto],'000') & [CodigoRiesgo])<>[CodigoUnico]));"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            m_IdRiesgo = .fields("IDRiesgo")
            m_CodigoUnico = .fields("CodUnico")
            If getcolRiesgosParaCambioCodigoCompleto Is Nothing Then
                Set getcolRiesgosParaCambioCodigoCompleto = New Scripting.Dictionary
                getcolRiesgosParaCambioCodigoCompleto.CompareMode = TextCompare
            End If
            If Not getcolRiesgosParaCambioCodigoCompleto.Exists(CStr(m_IdRiesgo)) Then
                getcolRiesgosParaCambioCodigoCompleto.Add m_IdRiesgo, m_CodigoUnico
            End If
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getcolRiesgosParaCambioCodigoCompleto ha devuelto el error: " & Err.description
    End If
End Function

Public Function RegularizarCodigoCompletoRiesgos(Optional ByRef p_Error As String) As String
    
    Dim m_SQL As String
    Dim m_RiesgosParaCambio As Scripting.Dictionary
    Dim m_IdRiesgo As Variant
    Dim m_CodUnico As String
    Dim m_Texto As String
    Dim i As Long
    
    
    On Error GoTo errores
    Set m_RiesgosParaCambio = getcolRiesgosParaCambioCodigoCompleto(p_Error)
    If m_RiesgosParaCambio Is Nothing Then
        RegularizarCodigoCompletoRiesgos = "Cambios Necesarios: 0"
        m_Texto = "Cambios Codigo Completo Riesgos: =0"
        VBA.DoEvents
        Debug.Print m_Texto
        VBA.DoEvents
        Exit Function
    End If
    m_Texto = "Cambios Codigo Completo Riesgos: " & m_RiesgosParaCambio.Count
    
    For Each m_IdRiesgo In m_RiesgosParaCambio
        m_CodUnico = m_RiesgosParaCambio(m_IdRiesgo)
        m_SQL = "UPDATE TbRiesgos SET CodigoUnico = '" & m_CodUnico & "' " & _
                "WHERE IDRiesgo=" & m_IdRiesgo & ";"
        getdb().Execute (m_SQL)
        i = i + 1
    Next
    m_Texto = m_Texto & vbNewLine & "Cambios Realizados: " & i
    VBA.DoEvents
    Debug.Print m_Texto
    VBA.DoEvents
    RegularizarCodigoCompletoRiesgos = m_Texto
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegularizarCodigoCompletoRiesgos ha devuelto el error: " & Err.description
    End If
    
End Function

Public Function RegularizarCambiosCompletosEnProyectos(Optional ByRef p_Error As String) As String
    
    Dim m_IDProyecto As Variant
    Dim m_Proyecto As Proyecto
    Dim m_Col As Scripting.Dictionary
    Dim i As Long
    
    On Error GoTo errores
    Set m_Col = getProyectosActivos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    i = 1
    For Each m_IDProyecto In m_Col
        Set m_Proyecto = m_Col(m_IDProyecto)
        VBA.DoEvents
        Debug.Print "Cambio para : " & m_Proyecto.Proyecto & " (" & i & " de " & m_Col.Count & ")"
        VBA.DoEvents
        With m_Proyecto
            .BorrarCambiosProyecto p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            .GrabarCambiosEnProyecto , p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End With
        i = i + 1
        Set m_Proyecto = Nothing
    Next
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegularizarCambiosCompletosEnProyectos ha devuelto el error: " & Err.description
    End If
    Debug.Print p_Error
End Function

Private Function getProyectosActivos( _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjProyecto As Proyecto
    
    On Error GoTo errores
    
    m_SQL = "SELECT * " & _
            "FROM TbProyectos " & _
            "WHERE FechaCierre Is Null;"
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
            If getProyectosActivos Is Nothing Then
                Set getProyectosActivos = New Scripting.Dictionary
                getProyectosActivos.CompareMode = TextCompare
            End If
            If Not getProyectosActivos.Exists(m_ObjProyecto.IDProyecto) Then
                getProyectosActivos.Add m_ObjProyecto.IDProyecto, m_ObjProyecto
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
        p_Error = "El método constructor.getProyectosActivos ha devuelto el error: " & Err.description
    End If
End Function
Private Function getRiesgosSinOrigen( _
                                        Optional ByRef p_Error As String _
                                        ) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE Origen Is Null Or Origen='';"
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
            If getRiesgosSinOrigen Is Nothing Then
                Set getRiesgosSinOrigen = New Scripting.Dictionary
                getRiesgosSinOrigen.CompareMode = TextCompare
            End If
            If Not getRiesgosSinOrigen.Exists(CStr(m_ObjRiesgo.IDRiesgo)) Then
                getRiesgosSinOrigen.Add CStr(m_ObjRiesgo.IDRiesgo), m_ObjRiesgo
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
        p_Error = "El método constructor.getRiesgosSinOrigen ha devuelto el error: " & Err.description
    End If
End Function

Public Function RegularizarOrigenesVacios(Optional ByRef p_Error As String) As String
    
    Dim m_Riesgo As riesgo
    Dim m_RiesgoAnterior As riesgo
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    Dim m_Origen As String
    Dim m_SQL As String
    Dim i As Long
    
    On Error GoTo errores
    Set m_Col = getRiesgosSinOrigen(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        VBA.DoEvents
        Debug.Print "RegularizarOrigenes 0 de 0"
        VBA.DoEvents
        Exit Function
    End If
    i = 1
    For Each m_ID In m_Col
        Set m_Riesgo = m_Col(m_ID)
        If Not m_Riesgo.RiesgoEdicionPrimera Is Nothing Then
            If Not m_Riesgo.RiesgoEdicionPrimera Is Nothing Then
                m_Origen = "Oferta"
                m_SQL = "UPDATE TbRiesgos SET Origen = '" & m_Origen & "' " & _
                        "WHERE IDRiesgo=" & m_ID & ";"
                getdb().Execute m_SQL
                VBA.DoEvents
                Debug.Print m_Riesgo.CodigoUnicoCalculado, m_Origen, i & " de " & m_Col.Count
                VBA.DoEvents
                GoTo siguiente
            End If
        End If
        Set m_RiesgoAnterior = m_Riesgo.RiesgoEdicionAnterior
        p_Error = m_Riesgo.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_RiesgoAnterior Is Nothing Then
            m_Origen = m_RiesgoAnterior.Origen
        Else
            If Not m_Riesgo.RiesgoExterno Is Nothing Then
                m_Origen = "Oferta"
            Else
                m_Origen = "Ejecución"
            End If
            
        End If
        m_SQL = "UPDATE TbRiesgos SET Origen = '" & m_Origen & "' " & _
                "WHERE IDRiesgo=" & m_ID & ";"
        getdb().Execute m_SQL
        VBA.DoEvents
        Debug.Print m_Riesgo.CodigoUnicoCalculado, m_Origen, i & " de " & m_Col.Count
        VBA.DoEvents
siguiente:
        i = i + 1

        Set m_Riesgo = Nothing
    Next
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegularizarOrigenesVacios ha devuelto el error: " & Err.description
    End If
    Debug.Print p_Error
End Function
Public Function RegularizarFechaMax(Optional ByRef p_Error As String) As String
    
    Dim m_Proyecto As Proyecto
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    Dim m_FechaMaxCalculada As String
    
    On Error GoTo errores
    Set m_Col = getProyectosActivos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        VBA.DoEvents
        Debug.Print "RegularizarFechaMax 0 de 0"
        VBA.DoEvents
        Exit Function
    End If
   
    For Each m_ID In m_Col
        Set m_Proyecto = m_Col(m_ID)
        m_FechaMaxCalculada = m_Proyecto.FechaMaxProximaPublicacionCalculada
        If m_Proyecto.FechaMaxProximaPublicacion <> m_FechaMaxCalculada Then
            Debug.Print m_Proyecto.NombreProyecto, m_Proyecto.FechaMaxProximaPublicacion, m_FechaMaxCalculada, m_Proyecto.ParaInformeAvisos
            SetFechaMaximaPublicacion m_Proyecto.IDProyecto, m_FechaMaxCalculada, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        

        Set m_Proyecto = Nothing
    Next
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegularizarFechaMax ha devuelto el error: " & Err.description
    End If
    Debug.Print p_Error
End Function

Private Function getcolRiesgoNC(Optional ByRef p_Error As String) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_IdRiesgo As String
    Dim m_IDNC As String
    Dim m_ParaNC As String
    Dim m_FechaDecision As String
    Dim m_Registro As String
    
    
    
    On Error GoTo errores
    m_SQL = "TbRiesgosNC"
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
            m_IDNC = Nz(.fields("IDNC").value, "")
            m_ParaNC = Nz(.fields("ParaNC").value, "")
            m_FechaDecision = Nz(.fields("FechaDecison").value, "")
            m_Registro = m_IdRiesgo & "|" & m_IDNC & "|" & m_ParaNC & "|" & m_FechaDecision
            
            If getcolRiesgoNC Is Nothing Then
                Set getcolRiesgoNC = New Scripting.Dictionary
                getcolRiesgoNC.CompareMode = TextCompare
            End If
            If Not getcolRiesgoNC.Exists(m_Registro) Then
                getcolRiesgoNC.Add m_Registro, m_Registro
            End If
            
            .MoveNext
        Loop
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getcolRiesgoNC ha devuelto el error: " & Err.description
    End If
    
End Function

Private Function getcolRiesgosMaterializados(Optional ByRef p_Error As String) As Scripting.Dictionary
    
    Dim m_Riesgo As riesgo
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    
    
    On Error GoTo errores
    'obtenemos riesgos en ediciones activas
    Set m_Col = Constructor.getRiesgosActivos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Riesgo = m_Col(m_ID)
        If IsDate(m_Riesgo.FechaMaterializado) Then
            If getcolRiesgosMaterializados Is Nothing Then
                Set getcolRiesgosMaterializados = New Scripting.Dictionary
                getcolRiesgosMaterializados.CompareMode = TextCompare
            End If
            If Not getcolRiesgosMaterializados.Exists(CStr(m_Riesgo.IDRiesgo)) Then
                getcolRiesgosMaterializados.Add CStr(m_Riesgo.IDRiesgo), m_Riesgo
            End If
        End If
        Set m_Riesgo = Nothing
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método constructor.getcolRiesgosMaterializados ha devuelto el error: " & Err.description
    End If
    
End Function

Public Function getcolRiesgoBorradosPorError(Optional ByRef p_Error As String) As Scripting.Dictionary
    
    Dim ColProyectosActivos As Scripting.Dictionary
    Dim m_IDProyecto As Variant
    Dim m_Proyecto As Proyecto
    Dim colEdiciones As Scripting.Dictionary
    Dim m_IDEdicion  As Variant
    Dim m_Edicion As Edicion
    Dim colRiesgos As Scripting.Dictionary
    Dim m_IdRiesgo As Variant
    Dim m_Riesgo As riesgo
    Dim ColRiesgosVistos As Scripting.Dictionary
    Dim m_BorradoPorError As EnumSiNo
    On Error GoTo errores
    
    Set ColProyectosActivos = getProyectosActivos(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If ColProyectosActivos Is Nothing Then
        Exit Function
    End If
    For Each m_IDProyecto In ColProyectosActivos
        Set m_Proyecto = ColProyectosActivos(m_IDProyecto)
        Set colEdiciones = m_Proyecto.colEdiciones
        If colEdiciones Is Nothing Then
            GoTo siguienteProyecto
        End If
        For Each m_IDEdicion In colEdiciones
            Set m_Edicion = colEdiciones(m_IDEdicion)
            Set colRiesgos = m_Edicion.colRiesgos
            If colRiesgos Is Nothing Then
                GoTo siguienteEdicion
            End If
            For Each m_IdRiesgo In colRiesgos
                Set m_Riesgo = colRiesgos(m_IdRiesgo)
                'If m_Riesgo.CodigoUnico = "012R005" Then Stop
                If Not ColRiesgosVistos Is Nothing Then
                    If ColRiesgosVistos.Exists(m_Riesgo.CodigoUnico) Then
                        GoTo siguienteRiesgo
                    End If
                End If
                m_BorradoPorError = m_Riesgo.RiesgoBorradoPorError
                If m_BorradoPorError = EnumSiNo.Sí Then
                    If getcolRiesgoBorradosPorError Is Nothing Then
                        Set getcolRiesgoBorradosPorError = New Scripting.Dictionary
                        getcolRiesgoBorradosPorError.CompareMode = TextCompare
                    End If
                    If Not getcolRiesgoBorradosPorError.Exists(CStr(m_Riesgo.IDRiesgo)) Then
                        getcolRiesgoBorradosPorError.Add CStr(m_Riesgo.IDRiesgo), m_Riesgo
                    End If
                End If
                
                If ColRiesgosVistos Is Nothing Then
                    Set ColRiesgosVistos = New Scripting.Dictionary
                    ColRiesgosVistos.CompareMode = TextCompare
                End If
                If Not ColRiesgosVistos.Exists(m_Riesgo.CodigoUnico) Then
                    ColRiesgosVistos.Add m_Riesgo.CodigoUnico, m_Riesgo.CodigoUnico
                End If
                
siguienteRiesgo:
                Set m_Riesgo = Nothing
            Next
siguienteEdicion:
            Set m_Edicion = Nothing
        Next
siguienteProyecto:
        Set m_Proyecto = Nothing
    Next
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getcolRiesgoBorradosPorError ha devuelto el error: " & Err.description
    End If
    
End Function

Public Function VerRiesgosEliminados(Optional ByRef p_Error As String) As String
    
    Dim m_ColRiesgosEliminados As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Riesgo As riesgo
    On Error GoTo errores
    Set m_ColRiesgosEliminados = getcolRiesgoBorradosPorError(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColRiesgosEliminados Is Nothing Then
        VerRiesgosEliminados = "Ninguno eliminado"
        Err.Raise 1000
    End If
    For Each m_ID In m_ColRiesgosEliminados
        Set m_Riesgo = m_ColRiesgosEliminados(m_ID)
        Debug.Print m_Riesgo.Edicion.Proyecto.Proyecto, m_Riesgo.CodigoRiesgo & "|" & "Ed. " & m_Riesgo.Edicion.Edicion
        
        Set m_Riesgo = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método VerRiesgosEliminados ha devuelto el error: " & Err.description
    End If
End Function

Public Function RegistrarMaterializacion( _
                                        p_RM As RiesgoMaterializacion, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                           
    ' Es cuando se materializa un riesgo o se desmaterializa se ha de rellenar
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_RiesgoMaterializacionEstadoAnterior As RiesgoMaterializacion
    
    
    On Error GoTo errores
    
    p_RM.Error = ""
    
   
    m_SQL = "TbRiesgosMaterializaciones"
    
    With p_RM
        .id = .IDCaculado
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not IsDate(.Fecha) Then
            .Fecha = Date
        End If
        Set m_RiesgoMaterializacionEstadoAnterior = .RiesgoMaterializacionEstadoAnterior
        If m_RiesgoMaterializacionEstadoAnterior Is Nothing Then
            If .EsMaterializacionCalcuado <> EnumSiNo.Sí Then
                Exit Function
            End If
        Else
            If .EsMaterializacionCalcuado = EnumSiNo.Sí Then
                If m_RiesgoMaterializacionEstadoAnterior.EsMaterializacionCalcuado = EnumSiNo.Sí Then
                    Exit Function
                End If
            Else
                If m_RiesgoMaterializacionEstadoAnterior.EsMaterializacionCalcuado = EnumSiNo.No Then
                    Exit Function
                End If
            End If
        End If
    End With
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .fields("ID") = p_RM.id
            .fields("IDProyecto") = p_RM.IDProyecto
            .fields("CodigoRiesgo") = p_RM.CodigoRiesgo
            .fields("IDEdicion") = p_RM.IDEdicion
            .fields("Fecha") = p_RM.Fecha
            .fields("EsMaterializacion") = p_RM.EsMaterializacion
            If p_RM.EsMaterializacion = "Sí" Then
                p_RM.Estado = "Materializado"
            End If
            If p_RM.Estado <> "" Then
                .fields("Estado") = p_RM.Estado
            End If
            
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    With p_RM
        If p_RM.Estado = "" Then
        .Estado = .riesgo.ESTADOCalculadoTexto
            m_SQL = "UPDATE TbRiesgosMaterializaciones SET TbRiesgosMaterializaciones.Estado = '" & .Estado & "' " & _
                    "WHERE ID=" & .id & ";"
            getdb().Execute m_SQL
        End If
    End With
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoMaterializacion.RegistrarMaterializacion ha devuelto el error: " & Err.description
    End If
End Function


