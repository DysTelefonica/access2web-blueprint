Attribute VB_Name = "Form_FormUsuario02HPS"
Option Compare Database
Option Explicit
Private m_Error As String

Private m_ObjColHPSParaBorrar As Scripting.Dictionary


Public Function HaHabidoCambiosPorTipoHPS( _
                                            p_TipoHPS As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Boolean
    
    Dim m_ObjActiva As Object
    Dim m_ObjAlInicio As Object
    On Error GoTo errores
    
    p_Error = ""
    RellenarHPSConDatosFormularioPorTIpoHPS p_TipoHPS, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If p_TipoHPS = "Nacional" Then
        Set m_ObjActiva = m_ObjHPSNACActiva
        Set m_ObjAlInicio = m_ObjHPSNACAlInicio
        
    ElseIf p_TipoHPS = "OTAN" Then
        Set m_ObjActiva = m_ObjHPSOTANActiva
        Set m_ObjAlInicio = m_ObjHPSOTANAlInicio
    ElseIf p_TipoHPS = "UE" Then
        Set m_ObjActiva = m_ObjHPSUEActiva
        Set m_ObjAlInicio = m_ObjHPSUEAlInicio
    ElseIf p_TipoHPS = "ESA" Then
        Set m_ObjActiva = m_ObjHPSESAActiva
        Set m_ObjAlInicio = m_ObjHPSESAAlInicio
    Else
        p_Error = "TIPOHPS desconocido"
        Err.Raise 1000
    End If
    If m_ObjActiva Is Nothing Then
        HaHabidoCambiosPorTipoHPS = False
        Exit Function
    End If
    If m_ObjAlInicio Is Nothing Then
        HaHabidoCambiosPorTipoHPS = True
        Exit Function
    End If
    With m_ObjAlInicio
        If .F_Concesion <> m_ObjActiva.F_Concesion Then
            HaHabidoCambiosPorTipoHPS = True
            Exit Function
        End If
        If .F_Caducidad <> m_ObjActiva.F_Caducidad Then
            HaHabidoCambiosPorTipoHPS = True
            Exit Function
        End If
        If .Grado <> m_ObjActiva.Grado Then
            HaHabidoCambiosPorTipoHPS = True
            Exit Function
        End If
        If .F_Baja <> m_ObjActiva.F_Baja Then
            HaHabidoCambiosPorTipoHPS = True
            Exit Function
        End If
        If .Renovacion <> m_ObjActiva.Renovacion Then
            HaHabidoCambiosPorTipoHPS = True
            Exit Function
        End If
        If .F_Solicitud <> m_ObjActiva.F_Solicitud Then
            HaHabidoCambiosPorTipoHPS = True
            Exit Function
        End If
        If .Especialidad <> m_ObjActiva.Especialidad Then
            HaHabidoCambiosPorTipoHPS = True
            Exit Function
        End If
        
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambiosPorTipoHPS ha producido el error: " & Err.Description
    End If

End Function
Private Sub ComandoGrabar_Click()
    
    Dim m_HaHabidoCambiosNAC As Boolean
    Dim m_HaHabidoCambiosOTAN As Boolean
    Dim m_HaHabidoCambiosESA As Boolean
    Dim m_HaHabidoCambiosUE As Boolean
    Dim blnEvaluarCambios As Boolean
    Dim m_ID As String
    Dim MotivoIrregularNacional As String
    Dim MotivoIrregularOTAN As String
    Dim MotivoIrregularESA As String
    Dim MotivoIrregularUE As String
    
    
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_ID = m_ObjUsuarioActivo.ID
    blnEvaluarCambios = False
    
    '   NACIONAL---------------
    '--------------------------
    If m_ObjColHPSParaBorrar Is Nothing Then
        blnEvaluarCambios = True
    Else
        If Not m_ObjColHPSParaBorrar.Exists("Nacional") Then
            blnEvaluarCambios = True
        End If
    End If
    If blnEvaluarCambios Then
        m_HaHabidoCambiosNAC = HaHabidoCambiosPorTipoHPS("Nacional", m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        MotivoIrregularNacional = Nz(Me.HPS_NAC_MOTIVO_IRREGULAR, "")
        If MotivoIrregularNacional <> "" Then
            m_Error = "No se pueden guardar los cambios al haber algún estado irregular"
            Err.Raise 1000
        End If
    End If
    
    '   OTAN---------------
    '--------------------------
    blnEvaluarCambios = False
    If m_ObjColHPSParaBorrar Is Nothing Then
        blnEvaluarCambios = True
    Else
        If Not m_ObjColHPSParaBorrar.Exists("OTAN") Then
            blnEvaluarCambios = True
        End If
    End If
    If blnEvaluarCambios Then
        m_HaHabidoCambiosOTAN = HaHabidoCambiosPorTipoHPS("OTAN", m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        MotivoIrregularNacional = Nz(Me.HPS_OTAN_MOTIVO_IRREGULAR, "")
        If MotivoIrregularNacional <> "" Then
            m_Error = "No se pueden guardar los cambios al haber algún estado irregular"
            Err.Raise 1000
        End If
    End If
    '   UE---------------
    '--------------------------
    blnEvaluarCambios = False
    If m_ObjColHPSParaBorrar Is Nothing Then
        blnEvaluarCambios = True
    Else
        If Not m_ObjColHPSParaBorrar.Exists("UE") Then
            blnEvaluarCambios = True
        End If
    End If
    If blnEvaluarCambios Then
        m_HaHabidoCambiosUE = HaHabidoCambiosPorTipoHPS("UE", m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        MotivoIrregularNacional = Nz(Me.HPS_UE_MOTIVO_IRREGULAR, "")
        If MotivoIrregularNacional <> "" Then
            m_Error = "No se pueden guardar los cambios al haber algún estado irregular"
            Err.Raise 1000
        End If
    End If
    '   ESA---------------
    '--------------------------
    blnEvaluarCambios = False
    If m_ObjColHPSParaBorrar Is Nothing Then
        blnEvaluarCambios = True
    Else
        If Not m_ObjColHPSParaBorrar.Exists("ESA") Then
            blnEvaluarCambios = True
        End If
    End If
    If blnEvaluarCambios Then
        m_HaHabidoCambiosESA = HaHabidoCambiosPorTipoHPS("ESA", m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        MotivoIrregularNacional = Nz(Me.HPS_ESA_MOTIVO_IRREGULAR, "")
        If MotivoIrregularNacional <> "" Then
            m_Error = "No se pueden guardar los cambios al haber algún estado irregular"
            Err.Raise 1000
        End If
    End If
    If Not m_ObjColHPSParaBorrar Is Nothing Then
        If m_ObjColHPSParaBorrar.Exists("Nacional") Then
            If m_ObjHPSNACActiva Is Nothing Then
                Set m_ObjHPSNACActiva = Constructor.getHPS("Nacional", m_ID)
                If Not m_ObjHPSNACActiva Is Nothing Then
                    m_ObjHPSNACActiva.Borrar m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                End If
                Set m_ObjHPSNACActiva = Nothing
                Set m_ObjHPSNACAlInicio = Nothing
            End If
        End If
        If m_ObjColHPSParaBorrar.Exists("OTAN") Then
            If m_ObjHPSOTANActiva Is Nothing Then
                Set m_ObjHPSOTANActiva = Constructor.getHPS("OTAN", m_ID)
                If Not m_ObjHPSOTANActiva Is Nothing Then
                    m_ObjHPSOTANActiva.Borrar m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                End If
                Set m_ObjHPSOTANActiva = Nothing
                Set m_ObjHPSOTANAlInicio = Nothing
            End If
        End If
        If m_ObjColHPSParaBorrar.Exists("UE") Then
            If m_ObjHPSUEActiva Is Nothing Then
                Set m_ObjHPSUEActiva = Constructor.getHPS("UE", m_ID)
                If Not m_ObjHPSUEActiva Is Nothing Then
                    m_ObjHPSUEActiva.Borrar m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                End If
                Set m_ObjHPSUEActiva = Nothing
                Set m_ObjHPSUEAlInicio = Nothing
            End If
        End If
        If m_ObjColHPSParaBorrar.Exists("ESA") Then
            If m_ObjHPSESAActiva Is Nothing Then
                Set m_ObjHPSESAActiva = Constructor.getHPS("ESA", m_ID)
                If Not m_ObjHPSESAActiva Is Nothing Then
                    m_ObjHPSESAActiva.Borrar m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                End If
                Set m_ObjHPSESAActiva = Nothing
                Set m_ObjHPSESAAlInicio = Nothing
            End If
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MsgBox("Datos borrados correctamente", vbInformation, "Datos borrados")
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_HaHabidoCambiosNAC Then
        If m_ObjHPSNACAlInicio Is Nothing Then
            m_ObjHPSNACActiva.ParaAlta = EnumSiNo.Sí
        Else
            m_ObjHPSNACActiva.ParaAlta = EnumSiNo.No
        End If
        m_ObjHPSNACActiva.Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If m_HaHabidoCambiosOTAN Then
        If m_ObjHPSOTANAlInicio Is Nothing Then
            m_ObjHPSOTANActiva.ParaAlta = EnumSiNo.Sí
        Else
            m_ObjHPSOTANActiva.ParaAlta = EnumSiNo.No
        End If
        m_ObjHPSOTANActiva.Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If m_HaHabidoCambiosUE Then
        If m_ObjHPSUEAlInicio Is Nothing Then
            m_ObjHPSUEActiva.ParaAlta = EnumSiNo.Sí
        Else
            m_ObjHPSUEActiva.ParaAlta = EnumSiNo.No
        End If
        m_ObjHPSUEActiva.Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If m_HaHabidoCambiosESA Then
        If m_ObjHPSESAAlInicio Is Nothing Then
            m_ObjHPSESAActiva.ParaAlta = EnumSiNo.Sí
        Else
            m_ObjHPSESAActiva.ParaAlta = EnumSiNo.No
        End If
        m_ObjHPSESAActiva.Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Set m_ObjUsuarioActivo = Constructor.getUsuarioHPS(m_ID, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    cache_usuario_regenerar m_ID, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    If m_HaHabidoCambiosNAC Or m_HaHabidoCambiosOTAN Or m_HaHabidoCambiosUE Or m_HaHabidoCambiosESA Then
        pregunta = MsgBox("Cambios registrados correctamente", vbInformation, "Registro")
    End If
    Form_FormUsuario01Lateral.lblDatosGenerales_Click
        
fin:
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
     DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarESA_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjColHPSParaBorrar Is Nothing Then
        Set m_ObjColHPSParaBorrar = New Scripting.Dictionary
        m_ObjColHPSParaBorrar.CompareMode = TextCompare
    End If
    If Not m_ObjColHPSParaBorrar.Exists("ESA") Then
        m_ObjColHPSParaBorrar.Add "ESA", "ESA"
    End If
    With Me
        .HPS_ESA_ESTADO = Null
        .HPS_ESA_F_Baja = Null
        .HPS_ESA_F_Caducidad = Null
        .HPS_ESA_F_Concesion = Null
        .HPS_ESA_F_Solicitud = Null
        .HPS_ESA_Grado = Null
        .HPS_ESA_MOTIVO_IRREGULAR = Null
        .HPS_ESA_Renovacion = Null
        .HPS_ESA_Especialidad = Null
    End With
    Set m_ObjHPSESAActiva = Nothing
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarESA_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarNacional_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjColHPSParaBorrar Is Nothing Then
        Set m_ObjColHPSParaBorrar = New Scripting.Dictionary
        m_ObjColHPSParaBorrar.CompareMode = TextCompare
    End If
    If Not m_ObjColHPSParaBorrar.Exists("Nacional") Then
        m_ObjColHPSParaBorrar.Add "Nacional", "Nacional"
    End If
    With Me
        .HPS_NAC_ESTADO = Null
        .HPS_NAC_F_Baja = Null
        .HPS_NAC_F_Caducidad = Null
        .HPS_NAC_F_Concesion = Null
        .HPS_NAC_F_Solicitud = Null
        .HPS_NAC_Grado = Null
        .HPS_NAC_MOTIVO_IRREGULAR = Null
        .HPS_NAC_Renovacion = Null
        .HPS_NAC_Especialidad = Null
    End With
    Set m_ObjHPSNACActiva = Nothing
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarNacional_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarOTAN_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjColHPSParaBorrar Is Nothing Then
        Set m_ObjColHPSParaBorrar = New Scripting.Dictionary
        m_ObjColHPSParaBorrar.CompareMode = TextCompare
    End If
    If Not m_ObjColHPSParaBorrar.Exists("OTAN") Then
        m_ObjColHPSParaBorrar.Add "OTAN", "OTAN"
    End If
    With Me
        .HPS_OTAN_ESTADO = Null
        .HPS_OTAN_F_Baja = Null
        .HPS_OTAN_F_Caducidad = Null
        .HPS_OTAN_F_Concesion = Null
        .HPS_OTAN_F_Solicitud = Null
        .HPS_OTAN_Grado = Null
        .HPS_OTAN_MOTIVO_IRREGULAR = Null
        .HPS_OTAN_Renovacion = Null
        .HPS_OTAN_Especialidad = Null
    End With
    Set m_ObjHPSOTANActiva = Nothing
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarOTAN_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarUE_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjColHPSParaBorrar Is Nothing Then
        Set m_ObjColHPSParaBorrar = New Scripting.Dictionary
        m_ObjColHPSParaBorrar.CompareMode = TextCompare
    End If
    If Not m_ObjColHPSParaBorrar.Exists("UE") Then
        m_ObjColHPSParaBorrar.Add "UE", "UE"
    End If
    With Me
        .HPS_UE_ESTADO = Null
        .HPS_UE_F_Baja = Null
        .HPS_UE_F_Caducidad = Null
        .HPS_UE_F_Concesion = Null
        .HPS_UE_F_Solicitud = Null
        .HPS_UE_Grado = Null
        .HPS_UE_MOTIVO_IRREGULAR = Null
        .HPS_UE_Renovacion = Null
        .HPS_UE_Especialidad = Null
    End With
    Set m_ObjHPSUEActiva = Nothing
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarUE_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoHPSESARenovacion_Click()
    If Me.HPS_ESA_Renovacion = "X" Then
        Me.HPS_ESA_Renovacion = Null
    Else
        Me.HPS_ESA_Renovacion = "X"
    End If
End Sub

Private Sub ComandoHPSNACRenovacion_Click()
    On Error GoTo errores
    If Me.HPS_NAC_Renovacion = "X" Then
        Me.HPS_NAC_Renovacion = Null
    Else
        Me.HPS_NAC_Renovacion = "X"
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar establecer la NO renovación se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoHPSOTANRenovacion_Click()
    If Me.HPS_OTAN_Renovacion = "X" Then
        Me.HPS_OTAN_Renovacion = Null
    Else
        Me.HPS_OTAN_Renovacion = "X"
    End If
End Sub

Private Sub ComandoHPSUERenovacion_Click()
    
    If Me.HPS_UE_Renovacion = "X" Then
        Me.HPS_UE_Renovacion = Null
    Else
        Me.HPS_UE_Renovacion = "X"
    End If

End Sub

Private Sub ComandoPasarAlResto_Click()

    
    On Error GoTo errores
    
    Me.HPS_OTAN_ESTADO = Null
    Me.HPS_OTAN_MOTIVO_IRREGULAR = Null
    Me.HPS_OTAN_F_Solicitud = Me.HPS_NAC_F_Solicitud
    Me.HPS_OTAN_Renovacion = Me.HPS_NAC_Renovacion
    Me.HPS_OTAN_F_Baja = Me.HPS_NAC_F_Baja
    Me.HPS_OTAN_F_Concesion = Me.HPS_NAC_F_Concesion
    Me.HPS_OTAN_F_Caducidad = Me.HPS_NAC_F_Caducidad
    Me.HPS_OTAN_Especialidad = Me.HPS_NAC_Especialidad
    
    Me.HPS_ESA_ESTADO = Null
    Me.HPS_ESA_MOTIVO_IRREGULAR = Null
    Me.HPS_ESA_F_Solicitud = Me.HPS_NAC_F_Solicitud
    Me.HPS_ESA_Renovacion = Me.HPS_NAC_Renovacion
    Me.HPS_ESA_F_Baja = Me.HPS_NAC_F_Baja
    Me.HPS_ESA_F_Concesion = Me.HPS_NAC_F_Concesion
    Me.HPS_ESA_F_Caducidad = Me.HPS_NAC_F_Caducidad
    Me.HPS_ESA_Especialidad = Me.HPS_NAC_Especialidad
    
    Me.HPS_UE_ESTADO = Null
    Me.HPS_UE_MOTIVO_IRREGULAR = Null
    Me.HPS_UE_F_Solicitud = Me.HPS_NAC_F_Solicitud
    Me.HPS_UE_Renovacion = Me.HPS_NAC_Renovacion
    Me.HPS_UE_F_Baja = Me.HPS_NAC_F_Baja
    Me.HPS_UE_F_Concesion = Me.HPS_NAC_F_Concesion
    Me.HPS_UE_F_Caducidad = Me.HPS_NAC_F_Caducidad
    Me.HPS_UE_Especialidad = Me.HPS_NAC_Especialidad
    
    If Me.HPS_NAC_Grado = "RESERVADO ( R )" Then
        Me.HPS_OTAN_Grado = "NATO SECRET ( NS )"
        Me.HPS_ESA_Grado = "ESA SECRET ( ESA-S )"
        Me.HPS_UE_Grado = "EU SECRET ( EU-S )"
    ElseIf Me.HPS_NAC_Grado = "SECRETO ( S )" Then
        Me.HPS_OTAN_Grado = "COSMIC TOP SECRET ( CTS )"
        Me.HPS_ESA_Grado = "ESA TOP SECRET ( ESA TS )"
        Me.HPS_UE_Grado = "EU TOP SECRET ( EU-TS )"
    ElseIf Me.HPS_NAC_Grado = "CONFIDENCIAL ( C )" Then
        Me.HPS_OTAN_Grado = "NATO CONFIDENTIAL ( NC )"
        Me.HPS_ESA_Grado = "ESA CONFIDENTIAL  ( ESA-C )"
        Me.HPS_UE_Grado = "EU CONFIDENTIAL ( EU-C )"
    End If
    RellenarHPSConDatosFormulario m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPasarAlResto_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub






Public Sub Form_Load()
    
    On Error GoTo errores
    m_Error = ""
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjUsuarioActivo Is Nothing Then
        m_Error = "No se ha abierto el formulario con los datos necesarios"
        Err.Raise 1000
    End If
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.AllowEdits = False
        Me.ComandoPasarAlResto.Enabled = False
        Me.ComandoEliminarNacional.Enabled = False
        Me.ComandoEliminarOTAN.Enabled = False
        Me.ComandoEliminarUE.Enabled = False
        Me.ComandoEliminarESA.Enabled = False
        Me.ComandoHPSNACRenovacion.Enabled = False
        Me.ComandoHPSOTANRenovacion.Enabled = False
        Me.ComandoHPSESARenovacion.Enabled = False
        Me.ComandoHPSUERenovacion.Enabled = False
    Else
        Me.AllowEdits = True
         Me.ComandoHPSNACRenovacion.Enabled = True
        Me.ComandoHPSOTANRenovacion.Enabled = True
        Me.ComandoHPSESARenovacion.Enabled = True
        Me.ComandoHPSUERenovacion.Enabled = True
    End If
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CopiarEstiloBoton Me.ComandoHPSNACRenovacion
    CopiarEstiloBoton Me.ComandoHPSOTANRenovacion
    CopiarEstiloBoton Me.ComandoHPSESARenovacion
    CopiarEstiloBoton Me.ComandoHPSUERenovacion
    CopiarEstiloBoton Me.ComandoGrabar
    CopiarEstiloBoton Me.ComandoPasarAlResto
   
    
    RellenarComboGrado "Nacional", Me.HPS_NAC_Grado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarComboGrado "OTAN", Me.HPS_OTAN_Grado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarComboGrado "UE", Me.HPS_UE_Grado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarComboGrado "ESA", Me.HPS_ESA_Grado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Application.TempVars("SoloLectura") <> "Sí" Then
        If Nz(Me.HPS_NAC_Grado = "CONFIDENCIAL ( C )") Or Nz(Me.HPS_NAC_Grado = "RESERVADO ( R )") Or Nz(Me.HPS_NAC_Grado, "") = "SECRETO ( S )" Then
            Me.ComandoPasarAlResto.Enabled = True
        Else
            Me.ComandoPasarAlResto.Enabled = False
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
    
End Sub

Private Function RellenarHPSConDatosFormularioPorTIpoHPS( _
                                                            p_TipoHPS As String, _
                                                            Optional ByRef p_Error As String _
                                                            ) As String
    
    
    Dim m_ParteNombreCampoHPS As String
    Dim m_NombreCampo As Variant
    Dim m_NombreCampoCompleto As String
    Dim m_ObjActiva As HPS
    Dim m_Valor As String
    Dim blnAlgunValor As Boolean
    
    On Error GoTo errores
    
    blnAlgunValor = False
    
    
    If p_TipoHPS = "Nacional" Then
        m_ParteNombreCampoHPS = "HPS_NAC_"
        Set m_ObjHPSNACActiva = Nothing
    ElseIf p_TipoHPS = "OTAN" Then
        m_ParteNombreCampoHPS = "HPS_OTAN_"
        Set m_ObjHPSOTANActiva = Nothing
    ElseIf p_TipoHPS = "UE" Then
        m_ParteNombreCampoHPS = "HPS_UE_"
        Set m_ObjHPSUEActiva = Nothing
    ElseIf p_TipoHPS = "ESA" Then
        m_ParteNombreCampoHPS = "HPS_ESA_"
        Set m_ObjHPSESAActiva = Nothing
    Else
        p_Error = "TIPOHPS desconocido"
        Err.Raise 1000
    End If
    Set m_ObjActiva = New HPS
    For Each m_NombreCampo In m_ObjActiva.ColCampos
        If CStr(m_NombreCampo) = "IDUsuario" Or CStr(m_NombreCampo) = "TipoHPS" Then
            GoTo siguiente
        End If
        m_NombreCampoCompleto = m_ParteNombreCampoHPS & m_NombreCampo
        m_Valor = Nz(Me.Controls(m_NombreCampoCompleto), "")
        If m_Valor <> "" Then
            blnAlgunValor = True
            Exit For
        End If
siguiente:
        
    Next
    If blnAlgunValor = False Then
        Exit Function
    End If
    m_ObjActiva.idUsuario = m_ObjUsuarioActivo.ID
    m_ObjActiva.tipoHps = p_TipoHPS
    For Each m_NombreCampo In m_ObjActiva.ColCampos
        If CStr(m_NombreCampo) = "IDUsuario" Or CStr(m_NombreCampo) = "TipoHPS" Then
            GoTo siguiente1
        End If
        m_NombreCampoCompleto = m_ParteNombreCampoHPS & m_NombreCampo
        m_Valor = Nz(Me.Controls(m_NombreCampoCompleto), "")
        m_ObjActiva.SetPropiedad m_NombreCampo, m_Valor, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
siguiente1:
    Next
    m_NombreCampoCompleto = m_ParteNombreCampoHPS & "ESTADO"
    Me.Controls(m_NombreCampoCompleto) = m_ObjActiva.Estado
    m_NombreCampoCompleto = m_ParteNombreCampoHPS & "MOTIVO_IRREGULAR"
    Me.Controls(m_NombreCampoCompleto) = m_ObjActiva.MOTIVO_IRREGULAR
    If p_TipoHPS = "Nacional" Then
        Set m_ObjHPSNACActiva = m_ObjActiva
    ElseIf p_TipoHPS = "OTAN" Then
        Set m_ObjHPSOTANActiva = m_ObjActiva
    ElseIf p_TipoHPS = "UE" Then
        Set m_ObjHPSUEActiva = m_ObjActiva
    ElseIf p_TipoHPS = "ESA" Then
        Set m_ObjHPSESAActiva = m_ObjActiva
    End If
    
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarHPSConDatosFormularioPorTIpoHPS  ha producido el error: " & Err.Description
    End If
End Function
Private Function RellenarHPSConDatosFormulario( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    
    
    
    On Error GoTo errores
    RellenarHPSConDatosFormularioPorTIpoHPS "Nacional", p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarHPSConDatosFormularioPorTIpoHPS "OTAN", p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarHPSConDatosFormularioPorTIpoHPS "UE", p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarHPSConDatosFormularioPorTIpoHPS "ESA", p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarHPSConDatosFormulario  ha producido el error: " & Err.Description
    End If
End Function
Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim m_ID As String
    Dim ctl As Control
    Dim m_NombreCampo As String
    Dim m_Valor As String
    
    On Error GoTo errores
    
    If m_ObjUsuarioActivo Is Nothing Then
        p_Error = "No se conoce el Usuario Activo"
        Err.Raise 1000
    End If
    m_ID = m_ObjUsuarioActivo.ID
    
    
    Set m_ObjHPSNACAlInicio = Nothing
    Set m_ObjHPSOTANAlInicio = Nothing
    Set m_ObjHPSESAAlInicio = Nothing
    Set m_ObjHPSUEAlInicio = Nothing
    
    
    
    Set m_ObjHPSNACAlInicio = Constructor.getHPS("Nacional", m_ID, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If Not m_ObjHPSNACAlInicio Is Nothing Then
        
        With m_ObjHPSNACAlInicio
            If .F_Solicitud <> "" Then
                Me.HPS_NAC_F_Solicitud = .F_Solicitud
            End If
            If .Grado <> "" Then
                Me.HPS_NAC_Grado = .Grado
            End If
            If .Renovacion <> "" Then
                If .Renovacion = "No" Then
                    Me.HPS_NAC_Renovacion = "X"
                Else
                    Me.HPS_NAC_Renovacion = Null
                End If
            End If
            If .F_Baja <> "" Then
                Me.HPS_NAC_F_Baja = .F_Baja
                
            End If
            If .F_Concesion <> "" Then
                Me.HPS_NAC_F_Concesion = .F_Concesion
            End If
            If .F_Caducidad <> "" Then
                Me.HPS_NAC_F_Caducidad = .F_Caducidad
            End If
            If .Estado <> "" Then
                Me.HPS_NAC_ESTADO = .Estado
            End If
            If .MOTIVO_IRREGULAR <> "" Then
                Me.HPS_NAC_MOTIVO_IRREGULAR = .MOTIVO_IRREGULAR
            End If
            If .Especialidad <> "" Then
                Me.HPS_NAC_Especialidad = .Especialidad
            End If
        End With
'        Me.ComandoEliminarNacional.Enabled = True
'        Me.ComandoPasarAlResto.Enabled = True
        
    Else
'        Me.ComandoEliminarNacional.Enabled = False
'        Me.ComandoPasarAlResto.Enabled = False
    End If
    
    Set m_ObjHPSOTANAlInicio = Constructor.getHPS("OTAN", m_ID, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjHPSOTANAlInicio Is Nothing Then
        With m_ObjHPSOTANAlInicio
            If .F_Solicitud <> "" Then
                Me.HPS_OTAN_F_Solicitud = .F_Solicitud
            End If
            If .Grado <> "" Then
                Me.HPS_OTAN_Grado = .Grado
            End If
            If .Renovacion <> "" Then
                If .Renovacion = "No" Then
                    Me.HPS_OTAN_Renovacion = "X"
                Else
                    Me.HPS_OTAN_Renovacion = Null
                End If
            End If
            If .F_Baja <> "" Then
                Me.HPS_OTAN_F_Baja = .F_Baja
                
            End If
            If .F_Concesion <> "" Then
                Me.HPS_OTAN_F_Concesion = .F_Concesion
            End If
            If .F_Caducidad <> "" Then
                Me.HPS_OTAN_F_Caducidad = .F_Caducidad
            End If
            If .Estado <> "" Then
                Me.HPS_OTAN_ESTADO = .Estado
            End If
            If .MOTIVO_IRREGULAR <> "" Then
                Me.HPS_OTAN_MOTIVO_IRREGULAR = .MOTIVO_IRREGULAR
            End If
            If .Especialidad <> "" Then
                Me.HPS_OTAN_Especialidad = .Especialidad
            End If
        End With
'        Me.ComandoEliminarOTAN.Enabled = True
    Else
'        Me.ComandoEliminarOTAN.Enabled = False
    End If
    Set m_ObjHPSUEAlInicio = Constructor.getHPS("UE", m_ID, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjHPSUEAlInicio Is Nothing Then
        With m_ObjHPSUEAlInicio
            If .F_Solicitud <> "" Then
                Me.HPS_UE_F_Solicitud = .F_Solicitud
            End If
            If .Grado <> "" Then
                Me.HPS_UE_Grado = .Grado
            End If
            If .Renovacion <> "" Then
                If .Renovacion = "No" Then
                    Me.HPS_UE_Renovacion = "X"
                Else
                    Me.HPS_UE_Renovacion = Null
                End If
            End If
            If .F_Baja <> "" Then
                Me.HPS_UE_F_Baja = .F_Baja
                
            End If
            If .F_Concesion <> "" Then
                Me.HPS_UE_F_Concesion = .F_Concesion
            End If
            If .F_Caducidad <> "" Then
                Me.HPS_UE_F_Caducidad = .F_Caducidad
            End If
            If .Estado <> "" Then
                Me.HPS_UE_ESTADO = .Estado
            End If
            If .MOTIVO_IRREGULAR <> "" Then
                Me.HPS_UE_MOTIVO_IRREGULAR = .MOTIVO_IRREGULAR
            End If
            If .Especialidad <> "" Then
                Me.HPS_UE_Especialidad = .Especialidad
            End If
        End With
'        Me.ComandoEliminarUE.Enabled = True
    Else
'        Me.ComandoEliminarUE.Enabled = False
    End If
    Set m_ObjHPSESAAlInicio = Constructor.getHPS("ESA", m_ID, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjHPSESAAlInicio Is Nothing Then
        With m_ObjHPSESAAlInicio
            If .F_Solicitud <> "" Then
                Me.HPS_ESA_F_Solicitud = .F_Solicitud
            End If
            If .Grado <> "" Then
                Me.HPS_ESA_Grado = .Grado
            End If
            If .Renovacion <> "" Then
                If .Renovacion = "No" Then
                    Me.HPS_ESA_Renovacion = "X"
                Else
                    Me.HPS_ESA_Renovacion = Null
                End If
            End If
            If .F_Baja <> "" Then
                Me.HPS_ESA_F_Baja = .F_Baja
                
            End If
            If .F_Concesion <> "" Then
                Me.HPS_ESA_F_Concesion = .F_Concesion
            End If
            If .F_Caducidad <> "" Then
                Me.HPS_ESA_F_Caducidad = .F_Caducidad
            End If
            If .Estado <> "" Then
                Me.HPS_ESA_ESTADO = .Estado
            End If
            If .MOTIVO_IRREGULAR <> "" Then
                Me.HPS_ESA_MOTIVO_IRREGULAR = .MOTIVO_IRREGULAR
            End If
            If .Especialidad <> "" Then
                Me.HPS_ESA_Especialidad = .Especialidad
            End If
        End With
'        Me.ComandoEliminarESA.Enabled = True
    Else
'        Me.ComandoEliminarESA.Enabled = False
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & Err.Description
    End If
End Function



Private Sub HPS_NAC_Grado_Change()
    If Application.TempVars("SoloLectura") <> "Sí" Then
        If Nz(Me.HPS_NAC_Grado, "") = "RESERVADO ( R )" Or Nz(Me.HPS_NAC_Grado, "") = "CONFIDENCIAL ( C )" Or Nz(Me.HPS_NAC_Grado, "") = "SECRETO ( S )" Then
            Me.ComandoPasarAlResto.Enabled = True
        Else
            Me.ComandoPasarAlResto.Enabled = False
        End If
    End If
End Sub
