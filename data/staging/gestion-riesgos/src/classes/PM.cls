VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDMitigacion As String
Public IDRiesgo As String
Public CodMitigacion As String
Public DisparadorDelPlan As String
Public FechaDeActivacion As String
Public FechaDesactivacion As String
Public Estado As String
Public NombreIcono As String

Private m_sIDMitigacionCalculado As String
Private m_ObjRiesgo As riesgo
Private m_ObjColAcciones As Scripting.Dictionary

Private m_eESTADOCalculado As EnumPlanEstado
Private m_eUsuarioConectadoAutorizado As EnumSiNo
Private m_sFechaActivacionCalculada As String
Private m_sFechaDesActivacionCalculada As String
Private m_eEsActivo As EnumSiNo
Private m_sCodigoMitigacionCalculado As String
Private m_objPMEdicionAnterior As PM
Private m_objPMEdicionSiguiente As PM
Private m_eTieneAcciones As EnumSiNo

Private m_objColAnteriores As Scripting.Dictionary

Private m_sESTADOCalculadoTexto As String
Private m_eNecesitaReplanificacion As EnumSiNo
Private m_sEstadoPlanParaEtiqueta As String
Private m_sFechaActivacionParaEtiqueta As String
Private m_sFechaDesactivacionParaEtiqueta As String
Private m_eBorrable As EnumSiNo
Private m_sNombreIconoCalculado As String
Private m_eAlgunaAccionParaReplanificar As EnumSiNo
Private m_objColCampos As Collection
Private m_objColCamposParaCambios As Collection
Public Error As String
Private m_Error As String

Public Property Get NombreIconoCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sNombreIconoCalculado = getNombreIconoPlanArbolDeRiesgos(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    NombreIconoCalculado = m_sNombreIconoCalculado
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método PM.NombreIconoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get Borrable() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eBorrable <> 0 Then
        Borrable = m_eBorrable
        Exit Property
    End If
    m_eBorrable = Me.riesgo.Borrable
    m_Error = Me.riesgo.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Borrable = m_eBorrable
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.Borrable ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NecesitaReplanificacion() As EnumSiNo
    
    Dim m_ID As Variant
    Dim m_PMAccion As PMAccion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.colAcciones Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eNecesitaReplanificacion = EnumSiNo.No
        NecesitaReplanificacion = m_eNecesitaReplanificacion
        Exit Property
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    For Each m_ID In Me.colAcciones
        Set m_PMAccion = Me.colAcciones(m_ID)
        With m_PMAccion
            If .NecesitaReplanificacion = EnumSiNo.Sí Then
                m_Error = .Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                m_eNecesitaReplanificacion = EnumSiNo.Sí
                NecesitaReplanificacion = m_eNecesitaReplanificacion
                Exit Property
            Else
                m_Error = .Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
        End With
        Set m_PMAccion = Nothing
    Next
   
    m_eNecesitaReplanificacion = EnumSiNo.No
    NecesitaReplanificacion = m_eNecesitaReplanificacion
    
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método PM.NecesitaReplanificacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property




Public Property Get IDMitigacionCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDMitigacionCalculado = DameID("TbRiesgosPlanMitigacionPpal", "IDMitigacion", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDMitigacionCalculado = m_sIDMitigacionCalculado
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método PM.IDMitigacionCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get riesgo() As riesgo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjRiesgo Is Nothing Then
        Set riesgo = m_ObjRiesgo
        Exit Property
    End If
    If Me.IDRiesgo = "" Then
        Exit Property
    End If
    Set m_ObjRiesgo = Constructor.getRiesgo(Me.IDRiesgo, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 100
    End If
    Set riesgo = m_ObjRiesgo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.Riesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set riesgo(ByVal sNewValue As riesgo)

    m_ObjRiesgo = sNewValue

End Property

Public Property Get colAcciones() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColAcciones Is Nothing Then
        Set colAcciones = m_ObjColAcciones
        Exit Property
    End If
    If Me.IDMitigacion = "" Then
        Exit Property
    End If
    Set m_ObjColAcciones = Constructor.getPMAcciones(Me.IDMitigacion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colAcciones = m_ObjColAcciones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.ColAcciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colAcciones(ByVal objNewValue As Variant)

    Set m_ObjColAcciones = objNewValue

End Property

Friend Sub InjectarColAcciones(p_colAcciones As Scripting.Dictionary)
    Set m_ObjColAcciones = p_colAcciones
End Sub

Public Property Get ESTADOCalculado() As EnumPlanEstado
        
    Dim m_IDAccion As Variant
    Dim m_ObjAccion As PMAccion
    Dim m_FechaFinPrevista As String
    Dim m_FechaInicio As String
    Dim m_FechaFinReal As String
    Dim m_FechaRealizado As String
    Dim m_NumeroAcciones As Integer
    Dim m_NumeroFechaFinReal As Integer
    Dim m_NumeroFechaFinPrevista As Integer
    Dim m_NumeroFechaInicio As Integer
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.TieneAcciones = EnumSiNo.No Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eESTADOCalculado = EnumPlanEstado.SinAcciones
        ESTADOCalculado = m_eESTADOCalculado
        Exit Property
    End If
    
    For Each m_IDAccion In Me.colAcciones
        If Nz(m_IDAccion, "") <> "" Then
            Set m_ObjAccion = Me.colAcciones(m_IDAccion)
            With m_ObjAccion
                
                If IsDate(.FechaInicio) Then
                    m_NumeroFechaInicio = m_NumeroFechaInicio + 1
                End If
                If IsDate(.FechaFinPrevista) Then
                    m_NumeroFechaFinPrevista = m_NumeroFechaFinPrevista + 1
                End If
                If IsDate(.FechaFinReal) Then
                    m_NumeroFechaFinReal = m_NumeroFechaFinReal + 1
                End If
            End With
            Set m_ObjAccion = Nothing
            m_NumeroAcciones = m_NumeroAcciones + 1
        End If
        
    Next
    If m_NumeroFechaFinReal = m_NumeroAcciones Then
        m_eESTADOCalculado = EnumPlanEstado.Finalizado
    Else
        If m_NumeroFechaInicio > 0 Then
            m_eESTADOCalculado = EnumPlanEstado.Activo
        Else
            If m_NumeroFechaFinPrevista > 0 Then
                m_eESTADOCalculado = EnumPlanEstado.Planificado
            Else
                m_eESTADOCalculado = EnumPlanEstado.Definido
            End If
        End If
    End If
    ESTADOCalculado = m_eESTADOCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.EstadoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Let ESTADOCalculado(ByVal eNewValue As EnumPlanEstado)

    m_eESTADOCalculado = eNewValue

End Property
Public Property Get FechaActivacionParaEtiqueta() As String
    
    Dim m_FechaActivacion As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_FechaActivacion = Me.FechaActivacionCalculada
    m_sFechaActivacionParaEtiqueta = "Primera Fecha Inicio: " & m_FechaActivacion
    
    
    FechaActivacionParaEtiqueta = m_sFechaActivacionParaEtiqueta
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.FechaActivacionParaEtiqueta ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaDesactivacionParaEtiqueta() As String
    
    Dim m_FechaDesactivacion As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_FechaDesactivacion = Me.FechaDesActivacionCalculada
    m_sFechaDesactivacionParaEtiqueta = "Última fecha de fin real: " & m_FechaDesactivacion
    
    
    FechaDesactivacionParaEtiqueta = m_sFechaDesactivacionParaEtiqueta
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.FechaDesactivacionParaEtiqueta ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get EstadoPlanParaEtiqueta() As String
    
    
    Dim m_EstadoTexto As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
    m_EstadoTexto = Me.ESTADOCalculadoTexto
    
    If m_EstadoTexto = "SinAcciones" Then
        m_sEstadoPlanParaEtiqueta = "Plan de Mitigación sin Acciones"
    Else
        m_sEstadoPlanParaEtiqueta = "Plan de Mitigación " & m_EstadoTexto
    End If
    
    
    
    EstadoPlanParaEtiqueta = m_sEstadoPlanParaEtiqueta
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.EstadoPlanParaEtiqueta ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get UsuarioConectadoAutorizado() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eUsuarioConectadoAutorizado <> 0 Then
        UsuarioConectadoAutorizado = m_eUsuarioConectadoAutorizado
        Exit Property
    End If
    m_eUsuarioConectadoAutorizado = UsuarioAutorizado(Me, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    UsuarioConectadoAutorizado = m_eUsuarioConectadoAutorizado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.UsuarioConectadoAutorizado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaActivacionCalculada() As String
    
    
    Dim m_IDAccion As Variant
    Dim m_ObjAccion As PMAccion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sFechaActivacionCalculada = ""
    If Me.colAcciones Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_IDAccion In Me.colAcciones.keys
        Set m_ObjAccion = Me.colAcciones(m_IDAccion)
        If IsDate(m_ObjAccion.FechaInicio) Then
            If IsDate(m_sFechaActivacionCalculada) Then
                If CDate(m_ObjAccion.FechaInicio) < CDate(m_sFechaActivacionCalculada) Then
                    m_sFechaActivacionCalculada = m_ObjAccion.FechaInicio
                End If
            Else
                m_sFechaActivacionCalculada = m_ObjAccion.FechaInicio
            End If
        End If
        Set m_ObjAccion = Nothing
    Next
    FechaActivacionCalculada = m_sFechaActivacionCalculada
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.FechaActivacionCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaDesActivacionCalculada() As String

    Dim m_IDAccion As Variant
    Dim m_ObjAccion As PMAccion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sFechaDesActivacionCalculada = "01/01/1900"
    
    If Me.colAcciones Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_IDAccion In Me.colAcciones.keys
        Set m_ObjAccion = Me.colAcciones(m_IDAccion)
        If Not IsDate(m_ObjAccion.FechaFinReal) Then
            Set m_ObjAccion = Nothing
            m_sFechaDesActivacionCalculada = ""
            FechaDesActivacionCalculada = m_sFechaDesActivacionCalculada
            Exit Property
        End If
        
        If IsDate(m_sFechaDesActivacionCalculada) Then
            If CDate(m_ObjAccion.FechaFinReal) > CDate(m_sFechaDesActivacionCalculada) Then
                m_sFechaDesActivacionCalculada = m_ObjAccion.FechaFinReal
            End If
        
        End If
        
        Set m_ObjAccion = Nothing
    Next
    If m_sFechaDesActivacionCalculada = "01/01/1900" Then
        m_sFechaDesActivacionCalculada = ""
    End If
    FechaDesActivacionCalculada = m_sFechaDesActivacionCalculada
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.FechaDesActivacionCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EsActivo() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eEsActivo <> 0 Then
        EsActivo = m_eEsActivo
        Exit Property
    End If
    m_eEsActivo = Me.riesgo.EsActivo
    m_Error = Me.riesgo.Error
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EsActivo = m_eEsActivo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.EsActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get CodigoMitigacionCalculado() As String
    
    Dim m_IDPM As Variant
    Dim m_ObjPM As PM
    Dim m_CodPM As String
    Dim m_intOrdinalMaxCod As Integer
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.riesgo Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se sabe a qué Riesgo pertenece"
        Err.Raise 1000
    End If
    If Not Me.riesgo.ColPMs Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_IDPM In Me.riesgo.ColPMs.keys
            Set m_ObjPM = Me.riesgo.ColPMs(m_IDPM)
            m_CodPM = m_ObjPM.CodMitigacion
            If IsNumeric(Right(m_CodPM, Len(m_CodPM) - 2)) Then
                If CInt(Right(m_CodPM, Len(m_CodPM) - 2)) > m_intOrdinalMaxCod Then
                    m_intOrdinalMaxCod = CInt(Right(m_CodPM, Len(m_CodPM) - 2))
                End If
            End If
            Set m_ObjPM = Nothing
        Next
    End If
    
    m_sCodigoMitigacionCalculado = "PM" & Format(m_intOrdinalMaxCod + 1, "00")
    
    CodigoMitigacionCalculado = m_sCodigoMitigacionCalculado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.CodigoMitigacionCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PMEdicionAnterior() As PM
    
    Dim m_ObjRiesgoAnterior As riesgo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objPMEdicionAnterior Is Nothing Then
        Set PMEdicionAnterior = m_objPMEdicionAnterior
        Exit Property
    End If
    If Me.riesgo Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido obtener el Riesgo del PM"
        Err.Raise 1000
    End If
    Set m_ObjRiesgoAnterior = Me.riesgo.RiesgoEdicionAnterior
    m_Error = Me.riesgo.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjRiesgoAnterior Is Nothing Then
        Exit Property
    End If
    Set m_objPMEdicionAnterior = m_ObjRiesgoAnterior.UltimoPM
    m_Error = m_ObjRiesgoAnterior.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set PMEdicionAnterior = m_objPMEdicionAnterior
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.PMEdicionAnterior ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get PMEdicionSiguiente() As PM
    
    Dim m_ObjRiesgoSiguiente As riesgo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objPMEdicionSiguiente Is Nothing Then
        Set PMEdicionSiguiente = m_objPMEdicionSiguiente
        Exit Property
    End If
    If Me.riesgo Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido obtener el Riesgo del PM"
        Err.Raise 1000
    End If
    Set m_ObjRiesgoSiguiente = Me.riesgo.RiesgoEdicionSiguiente
    m_Error = Me.riesgo.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjRiesgoSiguiente Is Nothing Then
        Exit Property
    End If
    Set m_objPMEdicionSiguiente = m_ObjRiesgoSiguiente.UltimoPM
    m_Error = m_ObjRiesgoSiguiente.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set PMEdicionSiguiente = m_objPMEdicionSiguiente
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.PMEdicionSiguiente ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ESTADOCalculadoTexto() As String
    
    Dim m_Estado As EnumPlanEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_Estado = Me.ESTADOCalculado
    If m_Estado = EnumPlanEstado.SinAcciones Then
        m_sESTADOCalculadoTexto = "SinAcciones"
    ElseIf m_Estado = EnumPlanEstado.Definido Then
        m_sESTADOCalculadoTexto = "Definido"
    ElseIf m_Estado = EnumPlanEstado.Planificado Then
        m_sESTADOCalculadoTexto = "Planificado"
    ElseIf m_Estado = EnumPlanEstado.Activo Then
        m_sESTADOCalculadoTexto = "Activo"
    ElseIf m_Estado = EnumPlanEstado.Finalizado Then
        m_sESTADOCalculadoTexto = "Finalizado"
    Else
        m_Error = "El estado del Plan no está reconocido"
        Err.Raise 1000
    End If
    ESTADOCalculadoTexto = m_sESTADOCalculadoTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.EstadoCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get TieneAcciones() As EnumSiNo
    Dim m_Acciones As Scripting.Dictionary
    
    Me.Error = ""
    m_Error = ""
    Set m_Acciones = Me.colAcciones
    If m_Acciones Is Nothing Then
        m_eTieneAcciones = EnumSiNo.No
    ElseIf m_Acciones.Count = 0 Then
        m_eTieneAcciones = EnumSiNo.No
    Else
        m_eTieneAcciones = EnumSiNo.Sí
    End If
    
    TieneAcciones = m_eTieneAcciones
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.TieneAcciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get AlgunaAccionParaReplanificar() As EnumSiNo
    Dim m_PMA As PMAccion
    Dim m_ID As Variant
    
    Me.Error = ""
    m_Error = ""
    If Me.colAcciones Is Nothing Then
        m_eAlgunaAccionParaReplanificar = EnumSiNo.No
        AlgunaAccionParaReplanificar = m_eAlgunaAccionParaReplanificar
        Exit Property
    End If
    For Each m_ID In Me.colAcciones
        Set m_PMA = Me.colAcciones(m_ID)
        If m_PMA.NecesitaReplanificacion = EnumSiNo.Sí Then
            m_eAlgunaAccionParaReplanificar = EnumSiNo.Sí
            AlgunaAccionParaReplanificar = m_eAlgunaAccionParaReplanificar
            Exit Property
        End If
        Set m_PMA = Nothing
    Next
    m_eAlgunaAccionParaReplanificar = EnumSiNo.No
    
    AlgunaAccionParaReplanificar = m_eAlgunaAccionParaReplanificar
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.AlgunaAccionParaReplanificar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get ColAnteriores() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_objColAnteriores = getPlanesIgualesEnEdcionesAnteriores(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnteriores = m_objColAnteriores
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.ColAnteriores ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColAnteriores(ByVal objNewValue As Variant)

    Set m_objColAnteriores = objNewValue

End Property



Public Function MotivoNoOK( _
                                Optional ByRef p_ObjPMAlInicio As PM, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    On Error GoTo errores
        
    With Me
        
        If .riesgo Is Nothing Then
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "No se conoce el riesgo"
            Exit Function
        End If
        If .riesgo.Estado = "" Then
            .riesgo.Estado = .riesgo.ESTADOCalculadoTexto
            p_Error = .riesgo.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If .riesgo.FichaRiesgoCompleta = EnumSiNo.No Then
            MotivoNoOK = "No se puede editar nin dar de alta un Plan de mitigación de un Riesgo al que le falta " & _
                        "completar los datos"
            Exit Function
        End If

        If .EsActivo <> EnumSiNo.Sí Then
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "El Riesgo no es activo"
            Exit Function
        End If
        If .DisparadorDelPlan = "" Then
            MotivoNoOK = "Obligatorio el disparador del plan"
            Exit Function
        End If
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function



Public Function Registrar( _
                            Optional ByRef p_ObjPMAlInicio As PM, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    m_MotivoNoOK = MotivoNoOK(p_ObjPMAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.CodMitigacion = "" Then
        Me.CodMitigacion = Me.CodigoMitigacionCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If Me.FechaDeActivacion = "" Then
        Me.FechaDeActivacion = Me.FechaActivacionCalculada
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If p_ObjPMAlInicio Is Nothing Then
        m_SQL = "TbRiesgosPlanMitigacionPpal"
        Me.IDMitigacion = Me.IDMitigacionCalculado
        
    Else
        m_SQL = "SELECT * FROM TbRiesgosPlanMitigacionPpal " & _
            "WHERE IDMitigacion=" & Me.IDMitigacion & ";"
        
    End If
    Me.Estado = Me.ESTADOCalculadoTexto
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not p_ObjPMAlInicio Is Nothing Then
            If .EOF Then
                p_Error = "No se ha encontrado el registro"
                Err.Raise 1000
            End If
        End If
        If p_ObjPMAlInicio Is Nothing Then
            .AddNew
                .Fields("IDMitigacion") = Me.IDMitigacion
                .Fields("IDRiesgo") = Me.IDRiesgo
                .Fields("CodMitigacion") = Me.CodMitigacion
                .Fields("DisparadorDelPlan") = Me.DisparadorDelPlan
                If IsDate(Me.FechaDeActivacion) Then
                    .Fields("FechaDeActivacion") = Me.FechaDeActivacion
                End If
                .Fields("Estado") = Me.Estado
                Me.NombreIcono = Me.NombreIconoCalculado
                If Me.NombreIcono <> "" Then
                    .Fields("NombreIcono") = Me.NombreIcono
                End If
            .Update
        Else
            .Edit
                .Fields("CodMitigacion") = Me.CodMitigacion
                .Fields("DisparadorDelPlan") = Me.DisparadorDelPlan
                If IsDate(Me.FechaDeActivacion) Then
                    .Fields("FechaDeActivacion") = Me.FechaDeActivacion
                Else
                    .Fields("FechaDeActivacion") = Null
                End If
                .Fields("Estado") = Me.Estado
                Me.NombreIcono = Me.NombreIconoCalculado
                If Me.NombreIcono <> "" Then
                    .Fields("NombreIcono") = Me.NombreIcono
                Else
                    .Fields("NombreIcono") = Null
                End If
            .Update
        
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set Me.riesgo.ColPMs = Nothing
    RegistrarActualizacionRiesgo p_Riesgo:=Me.riesgo, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionEdicion p_Edicion:=Me.riesgo.Edicion, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function


Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    
    
    
    '       1)- Sólo puede editar si proviene de una Edición donde el Autorizado (administradores,Calidad y JPs del Proyecto)
    '       2)- No puede provenir de un riesgo que pertenece a una Edición ya publicada
    
   
    On Error GoTo errores
    
    Me.Error = ""
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "La Edición a la que pertenece el Plan de mitigación no es activa"
        Err.Raise 1000
    End If
    If Me.UsuarioConectadoAutorizado = EnumSiNo.No Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
        Err.Raise 1000
    End If
        
        
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PM.ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                                
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "DELETE * FROM TbRiesgosPlanMitigacionDetalle " & _
            "WHERE IDMitigacion=" & Me.IDMitigacion & ";"
    getdb().Execute m_SQL
    
    m_SQL = "DELETE * FROM TbRiesgosPlanMitigacionPPal " & _
            "WHERE IDMitigacion=" & Me.IDMitigacion & ";"
    getdb().Execute m_SQL
    Set Me.riesgo.ColPMs = Nothing
    RegistrarActualizacionRiesgo p_Riesgo:=Me.riesgo, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionEdicion p_Edicion:=Me.riesgo.Edicion, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Borrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Cerrar( _
                        Optional ByRef p_Error As String _
                        ) As String
    '---------------------------------------------------------------------------------------------------------
    'si tiene fecha de inicio y no de fin se va a poner la de fin
    'va a ir acción a acción del plan de mitigación activo y la que tenga fecha de inicio le pone fecha fin
    '---------------------------------------------------------------------------------------------------------
   
    Dim m_SQL As String
    On Error GoTo errores
    Me.Error = ""
    m_SQL = "UPDATE TbRiesgosPlanMitigacionDetalle SET TbRiesgosPlanMitigacionDetalle.FechaFinReal = Now() " & _
            "WHERE (((TbRiesgosPlanMitigacionDetalle.IDMitigacion)=" & Me.IDMitigacion & "));"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getdb().Execute m_SQL
    Set Me.colAcciones = Nothing
   
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PM.Cerrar ha devuelto el error: " & Err.Description
    End If
                                                
End Function


Public Function GrabarNombreIcono( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    Me.NombreIcono = Me.NombreIconoCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "UPDATE TbRiesgosPlanMitigacionPpal SET NombreIcono = '" & Me.NombreIcono & "' " & _
            "WHERE IDMitigacion=" & Me.IDMitigacion & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PM.GrabarNombreIcono ha devuelto el error: " & Err.Description
    End If
End Function

Public Function GrabarEstado( _
                                    Optional p_Estado As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    If p_Estado = "" Then
        Me.Estado = Me.ESTADOCalculadoTexto
    Else
        Me.Estado = p_Estado
    End If
    
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "UPDATE TbRiesgosPlanMitigacionPpal SET Estado = '" & Me.Estado & "' " & _
            "WHERE IDMitigacion=" & Me.IDMitigacion & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PM.GrabarEstado ha devuelto el error: " & Err.Description
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
        .Add "IDMitigacion"
        .Add "IDRiesgo"
        .Add "CodMitigacion"
        .Add "DisparadorDelPlan"
        .Add "FechaDeActivacion"
        .Add "FechaDesactivacion"
        .Add "Estado"
        .Add "NombreIcono"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Property Get ColCamposParaCambios() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCamposParaCambios Is Nothing Then
        Set ColCamposParaCambios = m_objColCamposParaCambios
        Exit Property
    End If
    Set m_objColCamposParaCambios = New Collection
    With m_objColCamposParaCambios
        .Add "DisparadorDelPlan"
        .Add "FechaDeActivacion"
        .Add "FechaDesactivacion"
        .Add "Estado"
    End With
    Set ColCamposParaCambios = m_objColCamposParaCambios
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PM.ColCamposParaCambios ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDMitigacion" Then
        getPropiedad = Me.IDMitigacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDRiesgo" Then
        getPropiedad = Me.IDRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "CodMitigacion" Then
        getPropiedad = Me.CodMitigacion
        Exit Function
    End If
    If m_NombrePropiedad = "DisparadorDelPlan" Then
        getPropiedad = Me.DisparadorDelPlan
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDeActivacion" Then
        getPropiedad = Me.FechaDeActivacion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDesactivacion" Then
        getPropiedad = Me.FechaDesactivacion
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        getPropiedad = Me.Estado
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIcono" Then
        getPropiedad = Me.NombreIcono
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en PM.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PM.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDMitigacion" Then
        Me.IDMitigacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDRiesgo" Then
        Me.IDRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodMitigacion" Then
        Me.CodMitigacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DisparadorDelPlan" Then
        Me.DisparadorDelPlan = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDeActivacion" Then
        Me.FechaDeActivacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDesactivacion" Then
        Me.FechaDesactivacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        Me.Estado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIcono" Then
        Me.NombreIcono = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en PM.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PM.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function





