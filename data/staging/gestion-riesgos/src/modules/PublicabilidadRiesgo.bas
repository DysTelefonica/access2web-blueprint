Attribute VB_Name = "PublicabilidadRiesgo"
Option Compare Database
Option Explicit

Public Enum EnumPublicabilidadCheckEstado
    Cumple = 1
    NoCumple = 2
    NoAplica = 3
End Enum

Public Enum EnumPublicabilidadVeredicto
    Publicable = 1
    NoPublicable = 2
    NoAplica = 3
End Enum

Public Type tPublicabilidadRiesgoDatos
    CodigoRiesgo As String
    Descripcion As String
    EsEdicionActiva As EnumSiNo
    Estado As EnumRiesgoEstado
    Priorizacion As String
    RequiereRiesgoDeBiblioteca As EnumSiNo
    RiesgoParaRetipificar As EnumSiNo
    FechaRechazoAceptacionPorCalidad As String
    FechaRechazoRetiroPorCalidad As String
    FechaMaterializado As String
    IDPlanContingencia As String
    JustificacionAusenciaPC As String
    RiesgoAltoOMuyAlto As EnumSiNo
    TienePMs As EnumSiNo
    TodosPMFinalizados As EnumSiNo
    AlgunPMActivo As EnumSiNo
    AlgunPMSinAcciones As EnumSiNo
    AlgunPCSinAcciones As EnumSiNo
    TienePCs As EnumSiNo
    TodosPCFinalizados As EnumSiNo
    AlgunPCActivo As EnumSiNo
End Type

Public Function EvaluarPublicabilidadRiesgo( _
                                            ByRef p_Datos As tPublicabilidadRiesgoDatos, _
                                            ByRef p_Checks As Scripting.Dictionary, _
                                            ByRef p_Veredicto As EnumPublicabilidadVeredicto, _
                                            Optional ByRef p_Error As String _
                                            ) As EnumSiNo
    Dim m_Publicable As EnumSiNo
    Dim m_Index As Long
    Dim m_Estado As EnumRiesgoEstado
    Dim m_EnAceptacion As EnumSiNo
    Dim m_EnRetirada As EnumSiNo
    Dim m_AplicaBloque As Boolean
    Dim m_TextoDetalle As String

    On Error GoTo errores

    p_Error = ""
    m_Publicable = EnumSiNo.Sí
    p_Veredicto = EnumPublicabilidadVeredicto.Publicable
    m_Index = 1

    Set p_Checks = New Scripting.Dictionary
    p_Checks.CompareMode = TextCompare

    m_Estado = p_Datos.Estado
    m_EnAceptacion = RiesgoEnAceptacion(m_Estado)
    m_EnRetirada = RiesgoEnRetirada(m_Estado)

    If p_Datos.EsEdicionActiva = EnumSiNo.No Then
        p_Veredicto = EnumPublicabilidadVeredicto.NoAplica
        AgregarCheck p_Checks, m_Index, "edicion_activa", "Edición activa", EnumPublicabilidadCheckEstado.NoAplica, "Edicion no activa, riesgo ya publicado"
        AgregarChecksNoAplica p_Checks, m_Index
        EvaluarPublicabilidadRiesgo = EnumSiNo.Sí
        Exit Function
    End If

    AgregarCheck p_Checks, m_Index, "edicion_activa", "Edición activa", EnumPublicabilidadCheckEstado.Cumple

    If m_Estado = EnumRiesgoEstado.Aceptado Then
        AgregarCheck p_Checks, m_Index, "pm_con_acciones", "Plan de mitigación con acciones", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_con_acciones", "Plan de contingencia con acciones", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "datos_generales", "Datos generales cumplimentados", EnumPublicabilidadCheckEstado.NoAplica

        If IsNumeric(p_Datos.Priorizacion) Then
            AgregarCheck p_Checks, m_Index, "Priorizacion", "Priorización establecida", EnumPublicabilidadCheckEstado.Cumple
        Else
            AgregarCheck p_Checks, m_Index, "Priorizacion", "Priorización establecida", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        End If

        AgregarCheck p_Checks, m_Index, "Aceptación_calidad", "Aceptación aprobada por calidad", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "retirada_calidad", "Retirada aprobada por calidad", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "retipificacion", "Riesgo retipificado", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pm_activo_materializado", "Plan de mitigación activo (materializado)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_activo_materializado", "Plan de contingencia activo (materializado)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_asociado_materializacion", "Plan de contingencia asociado a la materialización", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_ausencia_justificada", "Ausencia de plan de contingencia justificada", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pm_activo_alto", "Plan de mitigación activo (alto/muy alto)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_definido_alto", "Plan de contingencia definido (alto/muy alto)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pm_definido_bajo", "Plan de mitigación definido (bajo/medio)", EnumPublicabilidadCheckEstado.NoAplica

        If m_Publicable = EnumSiNo.No Then
            p_Veredicto = EnumPublicabilidadVeredicto.NoPublicable
        End If
        EvaluarPublicabilidadRiesgo = m_Publicable
        Exit Function
    End If

    If p_Datos.TienePMs = EnumSiNo.Sí Then
        If p_Datos.AlgunPMSinAcciones = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "pm_con_acciones", "Plan de mitigación con acciones", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        Else
            AgregarCheck p_Checks, m_Index, "pm_con_acciones", "Plan de mitigación con acciones", EnumPublicabilidadCheckEstado.Cumple
        End If
    Else
        AgregarCheck p_Checks, m_Index, "pm_con_acciones", "Plan de mitigación con acciones", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If p_Datos.TienePCs = EnumSiNo.Sí Then
        If p_Datos.AlgunPCSinAcciones = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "pc_con_acciones", "Plan de contingencia con acciones", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        Else
            AgregarCheck p_Checks, m_Index, "pc_con_acciones", "Plan de contingencia con acciones", EnumPublicabilidadCheckEstado.Cumple
        End If
    Else
        AgregarCheck p_Checks, m_Index, "pc_con_acciones", "Plan de contingencia con acciones", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If m_Estado = EnumRiesgoEstado.Retirado Then
        AgregarCheck p_Checks, m_Index, "datos_generales", "Datos generales cumplimentados", EnumPublicabilidadCheckEstado.NoAplica
    ElseIf m_Estado <> EnumRiesgoEstado.Incompleto Then
        AgregarCheck p_Checks, m_Index, "datos_generales", "Datos generales cumplimentados", EnumPublicabilidadCheckEstado.Cumple
    Else
        AgregarCheck p_Checks, m_Index, "datos_generales", "Datos generales cumplimentados", EnumPublicabilidadCheckEstado.NoCumple
        m_Publicable = EnumSiNo.No
    End If

    If m_Estado = EnumRiesgoEstado.Retirado Then
        AgregarCheck p_Checks, m_Index, "Priorización", "Priorización establecida", EnumPublicabilidadCheckEstado.NoAplica
    ElseIf IsNumeric(p_Datos.Priorizacion) Then
        AgregarCheck p_Checks, m_Index, "Priorización", "Priorización establecida", EnumPublicabilidadCheckEstado.Cumple
    Else
        AgregarCheck p_Checks, m_Index, "Priorización", "Priorización establecida", EnumPublicabilidadCheckEstado.NoCumple
        m_Publicable = EnumSiNo.No
    End If

    If m_Estado = EnumRiesgoEstado.Aceptado Then
        AgregarCheck p_Checks, m_Index, "Aceptación_calidad", "Aceptación aprobada por calidad", EnumPublicabilidadCheckEstado.Cumple
    ElseIf m_EnAceptacion = EnumSiNo.Sí Then
        If p_Datos.FechaRechazoAceptacionPorCalidad <> "" Then
            m_TextoDetalle = "Rechazada por calidad"
        Else
            m_TextoDetalle = "Pendiente de evaluacion por calidad"
        End If
        AgregarCheck p_Checks, m_Index, "Aceptación_calidad", "Aceptación aprobada por calidad", EnumPublicabilidadCheckEstado.NoCumple, m_TextoDetalle
        m_Publicable = EnumSiNo.No
    Else
        AgregarCheck p_Checks, m_Index, "Aceptación_calidad", "Aceptación aprobada por calidad", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If m_Estado = EnumRiesgoEstado.Retirado Then
        AgregarCheck p_Checks, m_Index, "retirada_calidad", "Retirada aprobada por calidad", EnumPublicabilidadCheckEstado.Cumple
    ElseIf m_EnRetirada = EnumSiNo.Sí Then
        If p_Datos.FechaRechazoRetiroPorCalidad <> "" Then
            m_TextoDetalle = "Rechazada por calidad"
        Else
            m_TextoDetalle = "Pendiente de evaluacion por calidad"
        End If
        AgregarCheck p_Checks, m_Index, "retirada_calidad", "Retirada aprobada por calidad", EnumPublicabilidadCheckEstado.NoCumple, m_TextoDetalle
        m_Publicable = EnumSiNo.No
    Else
        AgregarCheck p_Checks, m_Index, "retirada_calidad", "Retirada aprobada por calidad", EnumPublicabilidadCheckEstado.NoAplica
    End If

    m_AplicaBloque = (m_Estado <> EnumRiesgoEstado.Aceptado And m_Estado <> EnumRiesgoEstado.Retirado)

    If m_AplicaBloque And p_Datos.RequiereRiesgoDeBiblioteca = EnumSiNo.Sí Then
        If p_Datos.RiesgoParaRetipificar = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "retipificacion", "Riesgo retipificado", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        Else
            AgregarCheck p_Checks, m_Index, "retipificacion", "Riesgo retipificado", EnumPublicabilidadCheckEstado.Cumple
        End If
    Else
        AgregarCheck p_Checks, m_Index, "retipificacion", "Riesgo retipificado", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If m_AplicaBloque And IsDate(p_Datos.FechaMaterializado) Then
        If p_Datos.AlgunPMActivo = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "pm_activo_materializado", "Plan de mitigación activo (materializado)", EnumPublicabilidadCheckEstado.Cumple
        Else
            AgregarCheck p_Checks, m_Index, "pm_activo_materializado", "Plan de mitigación activo (materializado)", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        End If
        If p_Datos.AlgunPCActivo = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "pc_activo_materializado", "Plan de contingencia activo (materializado)", EnumPublicabilidadCheckEstado.Cumple
        Else
            ' El plan de contingencia es OPCIONAL (decisión humana 2026-06-16, issue #50).
            ' Si no hay PC activo, no bloqueamos por este check: delegamos en pc_ausencia_justificada.
            AgregarCheck p_Checks, m_Index, "pc_activo_materializado", "Plan de contingencia activo (materializado)", EnumPublicabilidadCheckEstado.NoAplica
        End If

        ' Ausencia de plan de contingencia: si no hay PC activo, exigimos justificación.
        ' Sin justificación, el sistema bloquea la publicación.
        If p_Datos.AlgunPCActivo = EnumSiNo.Sí Then
            If Trim$(p_Datos.IDPlanContingencia) <> "" Then
                AgregarCheck p_Checks, m_Index, "pc_asociado_materializacion", "Plan de contingencia asociado a la materialización", EnumPublicabilidadCheckEstado.Cumple
            Else
                AgregarCheck p_Checks, m_Index, "pc_asociado_materializacion", "Plan de contingencia asociado a la materialización", EnumPublicabilidadCheckEstado.NoAplica
            End If
            AgregarCheck p_Checks, m_Index, "pc_ausencia_justificada", "Ausencia de plan de contingencia justificada", EnumPublicabilidadCheckEstado.NoAplica
        Else
            If Trim$(p_Datos.JustificacionAusenciaPC) <> "" Then
                AgregarCheck p_Checks, m_Index, "pc_asociado_materializacion", "Plan de contingencia asociado a la materialización", EnumPublicabilidadCheckEstado.NoAplica
                AgregarCheck p_Checks, m_Index, "pc_ausencia_justificada", "Ausencia de plan de contingencia justificada", EnumPublicabilidadCheckEstado.Cumple, "Justificada por el técnico: " & p_Datos.JustificacionAusenciaPC
            Else
                AgregarCheck p_Checks, m_Index, "pc_asociado_materializacion", "Plan de contingencia asociado a la materialización", EnumPublicabilidadCheckEstado.NoAplica
                AgregarCheck p_Checks, m_Index, "pc_ausencia_justificada", "Ausencia de plan de contingencia justificada", EnumPublicabilidadCheckEstado.NoCumple, "Materializado sin plan de contingencia activo y sin justificación"
                m_Publicable = EnumSiNo.No
            End If
        End If
    Else
        AgregarCheck p_Checks, m_Index, "pm_activo_materializado", "Plan de mitigación activo (materializado)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_activo_materializado", "Plan de contingencia activo (materializado)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_asociado_materializacion", "Plan de contingencia asociado a la materialización", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_ausencia_justificada", "Ausencia de plan de contingencia justificada", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If m_AplicaBloque And p_Datos.RiesgoAltoOMuyAlto = EnumSiNo.Sí And Not IsDate(p_Datos.FechaMaterializado) Then
        If p_Datos.AlgunPMActivo = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "pm_activo_alto", "Plan de mitigación activo (alto/muy alto)", EnumPublicabilidadCheckEstado.Cumple
        Else
            AgregarCheck p_Checks, m_Index, "pm_activo_alto", "Plan de mitigación activo (alto/muy alto)", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        End If
        If p_Datos.TienePCs = EnumSiNo.Sí And p_Datos.TodosPCFinalizados = EnumSiNo.No Then
            AgregarCheck p_Checks, m_Index, "pc_definido_alto", "Plan de contingencia definido (alto/muy alto)", EnumPublicabilidadCheckEstado.Cumple
        Else
            m_TextoDetalle = ""
            If p_Datos.TienePCs = EnumSiNo.No Then
                m_TextoDetalle = "Sin planes definidos"
            ElseIf p_Datos.TodosPCFinalizados = EnumSiNo.Sí Then
                m_TextoDetalle = "Todos los planes finalizados"
            End If
            AgregarCheck p_Checks, m_Index, "pc_definido_alto", "Plan de contingencia definido (alto/muy alto)", EnumPublicabilidadCheckEstado.NoCumple, m_TextoDetalle
            m_Publicable = EnumSiNo.No
        End If
    Else
        AgregarCheck p_Checks, m_Index, "pm_activo_alto", "Plan de mitigación activo (alto/muy alto)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_definido_alto", "Plan de contingencia definido (alto/muy alto)", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If m_AplicaBloque And p_Datos.RiesgoAltoOMuyAlto = EnumSiNo.No And Not IsDate(p_Datos.FechaMaterializado) Then
        If p_Datos.TienePMs = EnumSiNo.Sí And p_Datos.TodosPMFinalizados = EnumSiNo.No Then
            AgregarCheck p_Checks, m_Index, "pm_definido_bajo", "Plan de mitigación definido (bajo/medio)", EnumPublicabilidadCheckEstado.Cumple
        Else
            m_TextoDetalle = ""
            If p_Datos.TienePMs = EnumSiNo.No Then
                m_TextoDetalle = "Sin planes definidos"
            ElseIf p_Datos.TodosPMFinalizados = EnumSiNo.Sí Then
                m_TextoDetalle = "Todos los planes finalizados"
            End If
            AgregarCheck p_Checks, m_Index, "pm_definido_bajo", "Plan de mitigación definido (bajo/medio)", EnumPublicabilidadCheckEstado.NoCumple, m_TextoDetalle
            m_Publicable = EnumSiNo.No
        End If

    Else
        AgregarCheck p_Checks, m_Index, "pm_definido_bajo", "Plan de mitigación definido (bajo/medio)", EnumPublicabilidadCheckEstado.NoAplica

    End If

    If p_Veredicto = EnumPublicabilidadVeredicto.Publicable Then
        If m_Publicable = EnumSiNo.No Then
            p_Veredicto = EnumPublicabilidadVeredicto.NoPublicable
        End If
    End If

    EvaluarPublicabilidadRiesgo = m_Publicable
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo EvaluarPublicabilidadRiesgo ha devuelto el error: " & Err.Description
    End If
End Function

Public Function ConstruirDatosPublicabilidadRiesgo( _
                                                    ByRef p_Riesgo As riesgo, _
                                                    ByRef p_Datos As tPublicabilidadRiesgoDatos, _
                                                    Optional p_db As DAO.Database = Nothing, _
                                                    Optional ByRef p_Error As String _
                                                    ) As EnumSiNo
    Dim m_ColPMs As Scripting.Dictionary
    Dim m_ColPCs As Scripting.Dictionary
    Dim m_RiesgoMaterializadoUltimo As RiesgoMaterializacion
    Dim db As DAO.Database
    On Error GoTo errores

    p_Error = ""
    If p_Riesgo Is Nothing Then
        p_Error = "Se ha de indicar el riesgo"
        Err.Raise 1000
    End If

    p_Datos.CodigoRiesgo = p_Riesgo.CodigoRiesgo
    p_Datos.Descripcion = p_Riesgo.DescripcionParaLista
    If p_Riesgo.Edicion Is Nothing Then
        p_Error = "No se ha podido determinar la edicion"
        Err.Raise 1000
    End If
    p_Datos.EsEdicionActiva = p_Riesgo.Edicion.EsActivo
    p_Datos.Estado = p_Riesgo.EstadoEnum
    p_Datos.Priorizacion = p_Riesgo.Priorizacion
    p_Datos.RequiereRiesgoDeBiblioteca = p_Riesgo.RequiereRiesgoDeBibliotecaCalculado
    p_Datos.RiesgoParaRetipificar = p_Riesgo.RiesgoParaRetipificar
    p_Datos.FechaRechazoAceptacionPorCalidad = p_Riesgo.FechaRechazoAceptacionPorCalidad
    p_Datos.FechaRechazoRetiroPorCalidad = p_Riesgo.FechaRechazoRetiroPorCalidad
    p_Datos.FechaMaterializado = p_Riesgo.FechaMaterializado
    p_Datos.IDPlanContingencia = ""
    If IsDate(p_Datos.FechaMaterializado) Then
        Set m_RiesgoMaterializadoUltimo = p_Riesgo.RiesgoMaterializadoUltimo
        If p_Riesgo.Error <> "" Then
            p_Error = p_Riesgo.Error
            Err.Raise 1000
        End If
        If Not m_RiesgoMaterializadoUltimo Is Nothing Then
            p_Datos.IDPlanContingencia = m_RiesgoMaterializadoUltimo.IDPlanContingencia
        End If
    End If
    p_Datos.RiesgoAltoOMuyAlto = p_Riesgo.RiesgoAltoOMuyAlto

    If p_db Is Nothing Then
        Set m_ColPMs = p_Riesgo.ColPMs
    Else
        Set m_ColPMs = getPMsPub(p_Riesgo.IDRiesgo, , p_db, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If

    If m_ColPMs Is Nothing Then
        p_Datos.TienePMs = EnumSiNo.No
    Else
        p_Datos.TienePMs = EnumSiNo.Sí
    End If

    p_Datos.TodosPMFinalizados = CalcularTodosPlanesFinalizados(m_ColPMs, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    p_Datos.AlgunPMSinAcciones = CalcularPlanSinAcciones(m_ColPMs, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    p_Datos.AlgunPMActivo = CalcularAlgunPlanActivo(m_ColPMs, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    If p_db Is Nothing Then
        Set m_ColPCs = p_Riesgo.ColPCs
    Else
        Set m_ColPCs = getPCsPub(p_Riesgo.IDRiesgo, , p_db, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If

    If m_ColPCs Is Nothing Then
        p_Datos.TienePCs = EnumSiNo.No
    Else
        p_Datos.TienePCs = EnumSiNo.Sí
    End If

    p_Datos.TodosPCFinalizados = CalcularTodosPlanesFinalizados(m_ColPCs, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    p_Datos.AlgunPCSinAcciones = CalcularPlanSinAcciones(m_ColPCs, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    p_Datos.AlgunPCActivo = CalcularAlgunPlanActivo(m_ColPCs, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    ConstruirDatosPublicabilidadRiesgo = EnumSiNo.Sí
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo ConstruirDatosPublicabilidadRiesgo ha devuelto el error: " & Err.Description
    End If
    ConstruirDatosPublicabilidadRiesgo = EnumSiNo.No
End Function

Public Function TextoVeredictoPublicabilidad(ByVal p_Veredicto As EnumPublicabilidadVeredicto) As String

    Select Case p_Veredicto
        Case EnumPublicabilidadVeredicto.Publicable
            TextoVeredictoPublicabilidad = "Publicable"
        Case EnumPublicabilidadVeredicto.NoPublicable
            TextoVeredictoPublicabilidad = "No publicable"
        Case EnumPublicabilidadVeredicto.NoAplica
            TextoVeredictoPublicabilidad = "No aplica"
        Case Else
            TextoVeredictoPublicabilidad = "Desconocido"
    End Select

End Function

Public Function TextoEstadoPublicabilidad(ByVal p_Estado As EnumPublicabilidadCheckEstado) As String

    Select Case p_Estado
        Case EnumPublicabilidadCheckEstado.Cumple
            TextoEstadoPublicabilidad = "cumple"
        Case EnumPublicabilidadCheckEstado.NoCumple
            TextoEstadoPublicabilidad = "no_cumple"
        Case EnumPublicabilidadCheckEstado.NoAplica
            TextoEstadoPublicabilidad = "no_aplica"
        Case Else
            TextoEstadoPublicabilidad = "desconocido"
    End Select

End Function

Private Sub AgregarCheck( _
                        ByRef p_Checks As Scripting.Dictionary, _
                        ByRef p_Index As Long, _
                        ByVal p_Id As String, _
                        ByVal p_Texto As String, _
                        ByVal p_Estado As EnumPublicabilidadCheckEstado, _
                        Optional ByVal p_Detalle As String = "" _
                        )
    Dim m_Check As Scripting.Dictionary

    Set m_Check = New Scripting.Dictionary
    m_Check.CompareMode = TextCompare
    m_Check.Add "id", p_Id
    m_Check.Add "texto", p_Texto
    m_Check.Add "estado", p_Estado
    If p_Detalle <> "" Then
        m_Check.Add "detalle", p_Detalle
    End If

    p_Checks.Add CStr(p_Index), m_Check
    p_Index = p_Index + 1
End Sub

Private Sub AgregarChecksNoAplica( _
                                ByRef p_Checks As Scripting.Dictionary, _
                                ByRef p_Index As Long _
                                )
    Dim m_List As Variant
    Dim m_Item As Variant

    m_List = Array( _
        Array("datos_generales", "Datos generales cumplimentados"), _
        Array("Priorización", "Priorización establecida"), _
        Array("Aceptación_calidad", "Aceptación aprobada por calidad"), _
        Array("retirada_calidad", "Retirada aprobada por calidad"), _
        Array("pm_con_acciones", "Plan de mitigación con acciones"), _
        Array("pc_con_acciones", "Plan de contingencia con acciones"), _
        Array("retipificacion", "Riesgo retipificado"), _
        Array("pm_activo_materializado", "Plan de mitigación activo (materializado)"), _
        Array("pc_activo_materializado", "Plan de contingencia activo (materializado)"), _
        Array("pc_asociado_materializacion", "Plan de contingencia asociado a la materialización"), _
        Array("pc_ausencia_justificada", "Ausencia de plan de contingencia justificada"), _
        Array("pm_activo_alto", "Plan de mitigación activo (alto/muy alto)"), _
        Array("pc_definido_alto", "Plan de contingencia definido (alto/muy alto)"), _
        Array("pm_definido_bajo", "Plan de mitigación definido (bajo/medio)") _
    )

    For Each m_Item In m_List
        AgregarCheck p_Checks, p_Index, m_Item(0), m_Item(1), EnumPublicabilidadCheckEstado.NoAplica
    Next
End Sub

Private Function CalcularAlgunPlanActivo( _
                                        ByRef p_ColPlanes As Scripting.Dictionary, _
                                        Optional ByRef p_Error As String _
                                        ) As EnumSiNo
    Dim m_ID As Variant
    Dim m_Plan As Object
    Dim m_Acciones As Scripting.Dictionary

    On Error GoTo errores

    If p_ColPlanes Is Nothing Then
        CalcularAlgunPlanActivo = EnumSiNo.No
        Exit Function
    End If

    For Each m_ID In p_ColPlanes
        Set m_Plan = p_ColPlanes(m_ID)
        Set m_Acciones = m_Plan.colAcciones
        If Not m_Acciones Is Nothing Then
            If AlgunaAccionActiva(m_Acciones) = EnumSiNo.Sí Then
                CalcularAlgunPlanActivo = EnumSiNo.Sí
                Exit Function
            End If
        End If
        Set m_Plan = Nothing
    Next

    CalcularAlgunPlanActivo = EnumSiNo.No
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo CalcularAlgunPlanActivo ha devuelto el error: " & Err.Description
    End If
End Function

Private Function CalcularPlanSinAcciones( _
                                        ByRef p_ColPlanes As Scripting.Dictionary, _
                                        Optional ByRef p_Error As String _
                                        ) As EnumSiNo
    Dim m_ID As Variant
    Dim m_Plan As Object
    Dim m_Acciones As Scripting.Dictionary

    On Error GoTo errores

    p_Error = ""

    If p_ColPlanes Is Nothing Then
        CalcularPlanSinAcciones = EnumSiNo.No
        Exit Function
    End If

    For Each m_ID In p_ColPlanes
        Set m_Plan = p_ColPlanes(m_ID)
        Set m_Acciones = m_Plan.colAcciones
        If m_Acciones Is Nothing Or m_Acciones.Count = 0 Then
            CalcularPlanSinAcciones = EnumSiNo.Sí
            Exit Function
        End If
    Next

    CalcularPlanSinAcciones = EnumSiNo.No

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo CalcularPlanSinAcciones ha devuelto el error: " & Err.Description
    End If
End Function

Private Function CalcularTodosPlanesFinalizados( _
                                              ByRef p_ColPlanes As Scripting.Dictionary, _
                                              Optional ByRef p_Error As String _
                                              ) As EnumSiNo
    Dim m_ID As Variant
    Dim m_Plan As Object

    On Error GoTo errores

    p_Error = ""

    If p_ColPlanes Is Nothing Then
        CalcularTodosPlanesFinalizados = EnumSiNo.No
        Exit Function
    End If

    For Each m_ID In p_ColPlanes
        Set m_Plan = p_ColPlanes(m_ID)
        If m_Plan.ESTADOCalculado <> EnumPlanEstado.Finalizado Then
            CalcularTodosPlanesFinalizados = EnumSiNo.No
            Exit Function
        End If
    Next

    CalcularTodosPlanesFinalizados = EnumSiNo.Sí

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo CalcularTodosPlanesFinalizados ha devuelto el error: " & Err.Description
    End If
End Function

Private Function AlgunaAccionActiva( _
                                    ByRef p_Acciones As Scripting.Dictionary _
                                    ) As EnumSiNo
    Dim m_ID As Variant
    Dim m_Accion As Object

    If p_Acciones Is Nothing Then
        AlgunaAccionActiva = EnumSiNo.No
        Exit Function
    End If

    For Each m_ID In p_Acciones
        Set m_Accion = p_Acciones(m_ID)
        If IsDate(m_Accion.FechaFinPrevista) And Not IsDate(m_Accion.FechaFinReal) Then
            AlgunaAccionActiva = EnumSiNo.Sí
            Exit Function
        End If
        Set m_Accion = Nothing
    Next

    AlgunaAccionActiva = EnumSiNo.No
End Function

Public Function getPMsPub( _
                            p_IDRiesgo As String, _
                            Optional ByRef p_ParaLista As EnumSiNo, _
                            Optional p_db As DAO.Database = Nothing, _
                            Optional ByRef p_Error As String _
                            ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim db As DAO.Database
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_PM As PM
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
    If p_db Is Nothing Then
        Set db = getdb()
    Else
        Set db = p_db
    End If
    Set rcdDatos = db.OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_PM = New PM
            For Each m_Campo In m_PM.ColCampos
                m_PM.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getPMsPub Is Nothing Then
                Set getPMsPub = New Scripting.Dictionary
                getPMsPub.CompareMode = TextCompare
            End If
            If Not getPMsPub.Exists(m_PM.IDMitigacion) Then
                getPMsPub.Add m_PM.IDMitigacion, m_PM
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
        p_Error = "EL método constructor.getPMsPub ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getPCsPub( _
                        p_IDRiesgo As String, _
                        Optional ByRef p_ParaLista As EnumSiNo, _
                        Optional p_db As DAO.Database = Nothing, _
                        Optional ByRef p_Error As String _
                        ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim db As DAO.Database
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_PC As PC
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
    If p_db Is Nothing Then
        Set db = getdb()
    Else
        Set db = p_db
    End If
    Set rcdDatos = db.OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_PC = New PC
            For Each m_Campo In m_PC.ColCampos
                m_PC.SetPropiedad m_Campo, Nz(.Fields(m_Campo).value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getPCsPub Is Nothing Then
                Set getPCsPub = New Scripting.Dictionary
                getPCsPub.CompareMode = TextCompare
            End If
            If Not getPCsPub.Exists(m_PC.IDContingencia) Then
                getPCsPub.Add m_PC.IDContingencia, m_PC
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
        p_Error = "EL método constructor.getPCsPub ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

