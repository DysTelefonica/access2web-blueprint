VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Riesgo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public IDRiesgo As String
Public IDEdicion As String
Public CodigoUnico As String
Public CodigoRiesgo As String
Public FechaDetectado As String
Public DetectadoPor As String
Public EntidadDetecta As String
Public Plazo As String
Public Calidad As String
Public Coste As String
Public ImpactoGlobal As String
Public Vulnerabilidad As String
Public Valoracion As String
Public Mitigacion As String
Public Contingencia As String
Public RequierePlanContingencia As String
Public Descripcion As String
Public CausaRaiz As String
Public Estado As String
Public FechaEstado As String
Public Priorizacion As String
Public FechaMaterializado As String
Public FechaRetirado As String
Public FechaCerrado As String
Public FechaMitigacionAceptar As String
Public JustificacionAceptacionRiesgo As String
Public FechaJustificacionAceptacionRiesgo As String
Public FechaAprobacionAceptacionPorCalidad As String
Public FechaRechazoAceptacionPorCalidad As String
Public JustificacionRetiroRiesgo As String
Public FechaJustificacionRetiroRiesgo As String
Public FechaAprobacionRetiroPorCalidad As String
Public FechaRechazoRetiroPorCalidad As String
Public RequiereRiesgoDeBiblioteca As String
Public CodRiesgoBiblioteca As String
Public RiesgoPendienteRetipificacion As String
Public FechaRiesgoParaRetipificar As String
Public FechaRiesgoRetipificado As String
Public DiasSinRespuestaCalidadAceptacion As String
Public DiasSinRespuestaCalidadRetiro As String
Public DiasSinRespuestaCalidadRetipificacion As String
Public Origen As String
Public HayErrorEnRiesgo As String
Public Error As String
Public NombreNodoDesc As String
Public NombreNodoEstado As String
Public NombreIcono As String
Public ColorIcono As String

Private m_objColLineasInformePublicabilidad As Scripting.Dictionary

Private m_sURLInformeErroresRiesgo As String


Private m_sIDRiesgoCalculado As String
Private m_eRequiereRiesgoDeBibliotecaCalculado As EnumSiNo
Private m_sRequiereRiesgoDeBibliotecaCalculadoTexto As String
Private m_sImpactoGlobalCalculado As String
Private m_eRiesgoParaRetipificar As EnumSiNo
Private m_eESTADOCalculado As EnumRiesgoEstado

Private m_sRequierePlanContingenciaCalculado As EnumSiNo
Private m_sRequierePlanContingenciaCalculadoTexto As String
Private m_objEdicion As Edicion
Private m_objColPMs As Scripting.Dictionary
Private m_objColPCs As Scripting.Dictionary

Private m_ObjColAnexos As Scripting.Dictionary
Private m_ObjUltimoPM As PM
Private m_ObjUltimoPc As PC
Private m_ePublicable As EnumSiNo
Private m_bPublicableCalculado As Boolean
Private m_eHaNacidoenEstaEdicion As EnumSiNo
Private m_eBorrable As EnumSiNo
Private m_eProvieneDeRiesgoExterno As EnumSiNo
Private m_sValoracionCalculada As String
Private m_eRiesgoAltoOMuyAlto As EnumSiNo
Private m_sRiesgoAltoOMuyAltoTexto As String
Private m_ObjColMotivosNoPublicable As Scripting.Dictionary
Private m_eTienePriorizacionUltima As EnumSiNo
Private m_eTienePriorizacionPrimera As EnumSiNo
Private m_objRiesgoEdicionAnterior As riesgo
Private m_objRiesgoEdicionPrimera As riesgo
Private m_objRiesgoEdicionSiguiente As riesgo
Private m_eUsuarioConectadoAutorizado As EnumSiNo
Private m_eRiesgoDeEdicionActiva As EnumSiNo
Private m_sCodigoRiesgoCalculado As String
Private m_sCodigoUnicoCalculado As String
Private m_sPriorizacionCalculado As String
Private m_sFechaDetectadoCalculado As String
Private m_sFechaMaterializadoCalculado As String
Private m_ObjEdicionEnLaQueNace As Edicion
Private m_ObjEdicionUltimaEnLaQueExiste As Edicion
Private m_eEsActivo As EnumSiNo
Private m_objColRiesgosMismoCodigoUnico As Scripting.Dictionary
Private m_sESTADOCalculadoTexto As String
Private m_sContingenciaCalculada As String
Private m_ObjRiesgoExterno As RiesgoExterno
Private m_eTienePMs As EnumSiNo
Private m_eTienePCs As EnumSiNo
Private m_sRequisitos As String
Private m_ObjResponsableTelefonica As ExpedienteResponsable
Private m_sPriorizacionEdAnterior As String
Private m_objRiesgoEnEdicionActiva As riesgo


Public PteCalidad As String
Private m_sDescripcionParaLista As String
Private m_eAlgunPMActivo As EnumSiNo
Private m_eAlgunPCActivo As EnumSiNo
Private m_eAlgunPCSinAcciones As EnumSiNo
Private m_eAlgunPMSinAcciones As EnumSiNo

Private m_sParaInformeAvisos As String
Private m_eFichaRiesgoCompleta As EnumSiNo


Private m_ObjRiesgoAceptadoEnNacimiento As riesgo
Private m_ObjRiesgoRetiradoEnNacimiento As riesgo

Private m_sDiasRespuestaCalidadAceptacion As String 'desde FechaMitigacionAceptar hasta (FechaAprobacionAceptacionPorCalidad o FechaRechazoAceptacionPorCalidad)
Private m_sDiasRespuestaCalidadRetiro As String 'desde FechaJustificacionRetiroRiesgo hasta  (FechaAprobacionRetiroPorCalidad o FechaRechazoRetiroPorCalidad)
Private m_sDiasRespuestaCalidadRetipificacion As String 'desde FechaRiesgoParaRetipificar hasta FechaRiesgoRetipificado

Private m_sDiasSinRespuestaCalidadAceptacionCalculado As String
Private m_sDiasSinRespuestaCalidadRetiroCalculado As String
Private m_sDiasSinRespuestaCalidadRetipificacionCalculado As String

Private m_objRiesgoDeBiblioteca As RiesgoBiblioteca
Private m_objColMaterializaciones As Scripting.Dictionary

Private m_ObjRiesgoCuandoNace As riesgo
Private m_ObjRiesgoEnUltimaEdicion As riesgo
Private m_eRiesgoBorradoPorError As EnumSiNo
Private m_eRiesgoAlgunaVezMaterializado As EnumSiNo
Private m_eTodosPMFinalizados As EnumSiNo
Private m_eTodosPCFinalizados As EnumSiNo
Private m_eAlgunPMDefinido As EnumSiNo
Private m_objcolMaterializacionesPorDecidirNC As Scripting.Dictionary

Private m_ObjRiesgoMaterializadoUltimo As RiesgoMaterializacion
Private m_sErroresRiesgoTexto As String
Private m_eEstadoEnum As EnumRiesgoEstado


Private m_sHTMLRiesgoPorRetipificar As String
Private m_sHTMLRiesgoMaterializado As String
Private m_sNombreParaNodo As String

Private m_sHayErrorEnRiesgoCalculado As String
Private m_sNombreNodoDescCalculado As String
Private m_sNombreNodoEstadoCalculado As String
Private m_sNombreIconoCalculado As String
Private m_sColorIconoCalculado As String

Private m_objcolAccionesMitigacionSinCerrar As Scripting.Dictionary
Private m_objcolAccionesContingenciaSinCerrar As Scripting.Dictionary
Private m_eAlgunPMPlanificado As EnumSiNo
Private m_objColCampos As Collection
Private m_objColCamposParaCambios As Collection
Private m_Error As String

Public Property Get PriorizacionEdAnterior() As String

    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sPriorizacionEdAnterior <> "" Then
        PriorizacionEdAnterior = m_sPriorizacionEdAnterior
        Exit Property
    End If
    If Me.RiesgoEdicionAnterior Is Nothing Then
        m_sPriorizacionEdAnterior = "-"
    Else
        If IsNumeric(Me.RiesgoEdicionAnterior.Priorizacion) Then
            m_sPriorizacionEdAnterior = Me.RiesgoEdicionAnterior.Priorizacion
        Else
            m_sPriorizacionEdAnterior = "-"
        End If
        
    End If
    
    
    PriorizacionEdAnterior = m_sPriorizacionEdAnterior

    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.PriorizacionEdAnterior ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColorIconoCalculado() As String

    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If Me.EsActivo = EnumSiNo.No Then
        m_sColorIconoCalculado = "Negro"
        ColorIconoCalculado = m_sColorIconoCalculado
        Exit Property
    End If
    
    If Me.HayErrorEnRiesgo = "" Then
        Me.HayErrorEnRiesgoGrabar
        
    End If
    
    
    If Me.HayErrorEnRiesgo = "Sí" Then
         m_sColorIconoCalculado = "Rojo"
    Else
        m_sColorIconoCalculado = "Negro"
    End If
    
    ColorIconoCalculado = m_sColorIconoCalculado

    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColorIconoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NombreIconoCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sNombreIconoCalculado = getNombreIconoRiesgoArbolDeRiesgos(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    NombreIconoCalculado = m_sNombreIconoCalculado
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método Riesgo.NombreIconoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get NombreParaNodo() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sNombreParaNodo <> "" Then
        NombreParaNodo = m_sNombreParaNodo
        Exit Property
    End If
    m_sNombreParaNodo = Me.CodigoRiesgo & " ( " & Me.Edicion.Proyecto.NombreParaNodo & " )"
  
    
    NombreParaNodo = m_sNombreParaNodo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.NombreParaNodo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property





Public Property Get ParaInformeAvisos() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Edicion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_sParaInformeAvisos = Me.Edicion.ParaInformeAvisos
    ParaInformeAvisos = m_sParaInformeAvisos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ParaInformeAvisos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get AlgunPMActivo() As EnumSiNo
    
    Dim m_ID As Variant
    Dim m_PM As PM
    Dim m_PMEstado As EnumPlanEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.TienePMs = EnumSiNo.No Then
        m_eAlgunPMActivo = EnumSiNo.No
        AlgunPMActivo = m_eAlgunPMActivo
        Exit Property
    End If
    
    For Each m_ID In Me.ColPMs
        Set m_PM = Me.ColPMs(m_ID)
        m_PMEstado = m_PM.ESTADOCalculado
        If m_PMEstado = EnumPlanEstado.Activo Then
            m_eAlgunPMActivo = EnumSiNo.Sí
            AlgunPMActivo = m_eAlgunPMActivo
            Exit Property
        End If
        Set m_PM = Nothing
    Next
    m_eAlgunPMActivo = EnumSiNo.No
    AlgunPMActivo = m_eAlgunPMActivo
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.AlgunPMActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get AlgunPCActivo() As EnumSiNo
    
    Dim m_ID As Variant
    Dim m_PC As PC
    Dim m_PCEstado As EnumPlanEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.TienePCs = EnumSiNo.No Then
        m_eAlgunPCActivo = EnumSiNo.No
        AlgunPCActivo = m_eAlgunPCActivo
        Exit Property
    End If
    For Each m_ID In Me.ColPCs
        Set m_PC = Me.ColPCs(m_ID)
        m_PCEstado = m_PC.ESTADOCalculado
        If m_PCEstado = EnumPlanEstado.Activo Then
            m_eAlgunPCActivo = EnumSiNo.Sí
            AlgunPCActivo = m_eAlgunPCActivo
            Exit Property
        End If
        Set m_PC = Nothing
    Next
    m_eAlgunPCActivo = EnumSiNo.No
    AlgunPCActivo = m_eAlgunPCActivo
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.AlgunPCActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get AlgunPCSinAcciones() As EnumSiNo
    
    Dim m_ID As Variant
    Dim m_PC As PC
    Dim m_PCEstado As EnumPlanEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.TienePCs = EnumSiNo.No Then
        m_eAlgunPCSinAcciones = EnumSiNo.No
        AlgunPCSinAcciones = m_eAlgunPCSinAcciones
        Exit Property
    End If
    For Each m_ID In Me.ColPCs
        Set m_PC = Me.ColPCs(m_ID)
        m_PCEstado = m_PC.ESTADOCalculado
        If m_PCEstado = EnumPlanEstado.SinAcciones Then
            m_eAlgunPCSinAcciones = EnumSiNo.Sí
            AlgunPCSinAcciones = m_eAlgunPCSinAcciones
            Exit Property
        End If
        Set m_PC = Nothing
    Next
    m_eAlgunPCSinAcciones = EnumSiNo.No
    AlgunPCSinAcciones = m_eAlgunPCSinAcciones
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.AlgunPCSinAcciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get AlgunPMSinAcciones() As EnumSiNo
    
    Dim m_ID As Variant
    Dim m_PM As PM
    Dim m_PMEstado As EnumPlanEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.TienePMs = EnumSiNo.No Then
        m_eAlgunPMSinAcciones = EnumSiNo.No
        AlgunPMSinAcciones = m_eAlgunPMSinAcciones
        Exit Property
    End If
    For Each m_ID In Me.ColPMs
        Set m_PM = Me.ColPMs(m_ID)
        m_PMEstado = m_PM.ESTADOCalculado
        If m_PMEstado = EnumPlanEstado.SinAcciones Then
            m_eAlgunPMSinAcciones = EnumSiNo.Sí
            AlgunPMSinAcciones = m_eAlgunPMSinAcciones
            Exit Property
        End If
        Set m_PM = Nothing
    Next
    m_eAlgunPMSinAcciones = EnumSiNo.No
    AlgunPMSinAcciones = m_eAlgunPMSinAcciones
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.AlgunPMSinAcciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DescripcionParaLista() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sDescripcionParaLista <> "" Then
        DescripcionParaLista = m_sDescripcionParaLista
        Exit Property
    End If
    If Me.Descripcion = "" Then
        Exit Property
    End If
    If Me.CausaRaiz = "" Then
        m_sDescripcionParaLista = Me.Descripcion
    Else
        m_sDescripcionParaLista = Me.Descripcion & "." & " " & Me.CausaRaiz
    End If
    
    DescripcionParaLista = m_sDescripcionParaLista
    
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.PMNecesitaReplanificacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get IDRiesgoCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDRiesgoCalculado = DameID("TbRiesgos", "IDRiesgo", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    IDRiesgoCalculado = m_sIDRiesgoCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.IDRiesgoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get RequiereRiesgoDeBibliotecaCalculado() As EnumSiNo
    
    On Error GoTo errores
    
    m_Error = ""
    Me.Error = ""
    If m_eRequiereRiesgoDeBibliotecaCalculado <> 0 Then
        RequiereRiesgoDeBibliotecaCalculado = m_eRequiereRiesgoDeBibliotecaCalculado
        Exit Property
    End If
    If Me.Edicion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If Me.Edicion.Proyecto Is Nothing Then
        m_Error = Me.Edicion.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    m_eRequiereRiesgoDeBibliotecaCalculado = Me.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado
    m_Error = Me.Edicion.Proyecto.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RequiereRiesgoDeBibliotecaCalculado = m_eRequiereRiesgoDeBibliotecaCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RequiereRiesgoDeBibliotecaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let RequiereRiesgoDeBibliotecaCalculado(ByVal sNewValue As EnumSiNo)

    m_eRequiereRiesgoDeBibliotecaCalculado = sNewValue

End Property
Public Property Get RequiereRiesgoDeBibliotecaCalculadoTexto() As String
    
    On Error GoTo errores
    
    m_Error = ""
    Me.Error = ""
    If m_sRequiereRiesgoDeBibliotecaCalculadoTexto <> "" Then
        RequiereRiesgoDeBibliotecaCalculadoTexto = m_sRequiereRiesgoDeBibliotecaCalculadoTexto
        Exit Property
    End If
    If Me.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
        m_sRequiereRiesgoDeBibliotecaCalculadoTexto = "Sí"
    ElseIf Me.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
        m_sRequiereRiesgoDeBibliotecaCalculadoTexto = "No"
    End If
    
    RequiereRiesgoDeBibliotecaCalculadoTexto = m_sRequiereRiesgoDeBibliotecaCalculadoTexto
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RequiereRiesgoDeBibliotecaCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property
Public Property Get RiesgoParaRetipificar() As EnumSiNo
    
    Dim m_RiesgoBiblioteca As RiesgoBiblioteca
    On Error GoTo errores
    
    m_Error = ""
    Me.Error = ""
    Set m_RiesgoBiblioteca = Me.RiesgoDeBiblioteca
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RiesgoBiblioteca Is Nothing Then
        m_eRiesgoParaRetipificar = EnumSiNo.No
        RiesgoParaRetipificar = m_eRiesgoParaRetipificar
        Exit Property
    End If
    m_eRiesgoParaRetipificar = m_RiesgoBiblioteca.RiesgoParaRetipificar
    m_Error = m_RiesgoBiblioteca.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RiesgoParaRetipificar = m_eRiesgoParaRetipificar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoParaRetipificar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let RiesgoParaRetipificar(ByVal sNewValue As EnumSiNo)

    m_eRiesgoParaRetipificar = sNewValue

End Property

Public Property Get AlgunPMPlanificado() As EnumSiNo
    
    Dim m_ID As Variant
    Dim m_PM As PM
    Dim m_PMEstado As EnumPlanEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.TienePMs = EnumSiNo.No Then
        m_eAlgunPMPlanificado = EnumSiNo.No
        AlgunPMPlanificado = m_eAlgunPMPlanificado
        Exit Property
    End If
    For Each m_ID In Me.ColPMs
        Set m_PM = Me.ColPMs(m_ID)
        m_PMEstado = m_PM.ESTADOCalculado
        If m_PMEstado = EnumPlanEstado.Planificado Then
            m_eAlgunPMPlanificado = EnumSiNo.Sí
            AlgunPMPlanificado = m_eAlgunPMPlanificado
            Exit Property
        End If
        Set m_PM = Nothing
    Next
    m_eAlgunPMPlanificado = EnumSiNo.No
    AlgunPMPlanificado = m_eAlgunPMPlanificado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.AlgunPMPlanificado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ESTADOCalculado() As EnumRiesgoEstado
    
    ' al haber varios planes de mitigación ahora podría el riesgo tener varios estados en función de cada uno de los
    'PM
    ' por ello que tomo este criterio
    ' si hay algún plan activo -->Estado activo
    ' si hay algún plan planificado -->Estado Planificado
    ' si hay algún plan definido --> Estado definido
    ' si hay algún plan incompleto -->Estado Incompleto
    ' si todos los planes están finalizados o no tiene -->Detectado
    
    
    Dim m_EsUltimaEdicion As EnumSiNo
    
    Dim m_ColPMs As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Mitigacion = "Aceptar" Then
        If Me.JustificacionAceptacionRiesgo = "" Then
            m_eESTADOCalculado = EnumRiesgoEstado.AceptadoSinJustificar
            ESTADOCalculado = m_eESTADOCalculado
            Exit Property
        End If
        If Not IsDate(Me.FechaAprobacionAceptacionPorCalidad) And Not IsDate(Me.FechaRechazoAceptacionPorCalidad) Then
            m_eESTADOCalculado = EnumRiesgoEstado.AceptadoSinVisar
            ESTADOCalculado = m_eESTADOCalculado
            Exit Property
        End If
        If IsDate(Me.FechaAprobacionAceptacionPorCalidad) Then
            m_eESTADOCalculado = EnumRiesgoEstado.Aceptado
        Else
            m_eESTADOCalculado = EnumRiesgoEstado.AceptadoRechazado
        End If
        ESTADOCalculado = m_eESTADOCalculado
        Exit Property
    End If
    If IsDate(Me.FechaRetirado) Then
        If Me.JustificacionRetiroRiesgo = "" Then
            m_eESTADOCalculado = EnumRiesgoEstado.RetiradoSinJustificar
            ESTADOCalculado = m_eESTADOCalculado
            Exit Property
        End If
        If Not IsDate(Me.FechaAprobacionRetiroPorCalidad) And Not IsDate(Me.FechaRechazoRetiroPorCalidad) Then
            m_eESTADOCalculado = EnumRiesgoEstado.RetiradoSinVisar
            ESTADOCalculado = m_eESTADOCalculado
            Exit Property
        End If
        If IsDate(Me.FechaAprobacionRetiroPorCalidad) Then
            m_eESTADOCalculado = EnumRiesgoEstado.Retirado
        Else
            m_eESTADOCalculado = EnumRiesgoEstado.RetiradoRechazado
        End If
        ESTADOCalculado = m_eESTADOCalculado
        Exit Property
    End If
    If IsDate(Me.FechaMaterializado) Then
        m_eESTADOCalculado = EnumRiesgoEstado.Materializado
        ESTADOCalculado = m_eESTADOCalculado
        Exit Property
    End If
    
    If Me.Plazo = "" Or Me.Coste = "" Or Me.Calidad = "" Or Me.ImpactoGlobal = "" Or Me.Valoracion = "" Or _
        Me.Vulnerabilidad = "" Or Me.Mitigacion = "" Or Me.FechaDetectado = "" Or Me.EntidadDetecta = "" Or _
        Me.DetectadoPor = "" Or Me.Descripcion = "" Or Me.Priorizacion = "" Then
        m_eESTADOCalculado = EnumRiesgoEstado.Incompleto
        ESTADOCalculado = m_eESTADOCalculado
        Exit Property
    End If
    
    If Me.Edicion.Proyecto.FechaCierre <> "" Then
        If Me.Edicion.Proyecto.EdicionUltimaNoActiva.IDEdicion = Me.IDEdicion Then
            m_EsUltimaEdicion = EnumSiNo.Sí
        Else
            m_EsUltimaEdicion = EnumSiNo.No
           
        End If
    Else
        m_EsUltimaEdicion = EnumSiNo.No
    End If
    'si estamos en la última edición y ya no está activo el proyecto ha de estar cerrado
    If m_EsUltimaEdicion = EnumSiNo.Sí Then
        m_eESTADOCalculado = EnumRiesgoEstado.Cerrado
        ESTADOCalculado = m_eESTADOCalculado
        Exit Property
    End If
    If Me.TienePMs = EnumSiNo.No Then
        m_eESTADOCalculado = EnumRiesgoEstado.Detectado
        ESTADOCalculado = m_eESTADOCalculado
        Exit Property
    End If
    
    If Me.TodosPMFinalizados = EnumSiNo.Sí Then
        m_eESTADOCalculado = EnumRiesgoEstado.Detectado
        ESTADOCalculado = m_eESTADOCalculado
        Exit Property
        
    End If
    If Me.AlgunPMActivo = EnumSiNo.Sí Then
        m_eESTADOCalculado = EnumRiesgoEstado.Activo
        ESTADOCalculado = m_eESTADOCalculado
        Exit Property
    End If
    If Me.AlgunPMPlanificado = EnumSiNo.Sí Then
        m_eESTADOCalculado = EnumRiesgoEstado.Planificado
        ESTADOCalculado = m_eESTADOCalculado
        Exit Property
    End If
    m_eESTADOCalculado = EnumRiesgoEstado.Detectado
    ESTADOCalculado = m_eESTADOCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.EstadoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Let ESTADOCalculado(ByVal m_eVal As EnumRiesgoEstado)

    m_eESTADOCalculado = m_eVal

End Property

Public Property Get TodosPMFinalizados() As EnumSiNo
    
    Dim m_ID As Variant
    Dim m_PM As PM
    Dim m_PMEstado As EnumPlanEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.TienePMs = EnumSiNo.No Then
        m_eTodosPMFinalizados = EnumSiNo.No
        TodosPMFinalizados = m_eTodosPMFinalizados
        Exit Property
    End If
    For Each m_ID In Me.ColPMs
        Set m_PM = Me.ColPMs(m_ID)
        m_PMEstado = m_PM.ESTADOCalculado
        If m_PMEstado <> EnumPlanEstado.Finalizado Then
            m_eTodosPMFinalizados = EnumSiNo.No
            TodosPMFinalizados = m_eTodosPMFinalizados
            Exit Property
        End If
        Set m_PM = Nothing
    Next
    m_eTodosPMFinalizados = EnumSiNo.Sí
    TodosPMFinalizados = m_eTodosPMFinalizados
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.TodosPMFinalizados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get TodosPCFinalizados() As EnumSiNo
    
    Dim m_ID As Variant
    Dim m_PC As PC
    Dim m_PCEstado As EnumPlanEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.TienePCs = EnumSiNo.No Then
        m_eTodosPCFinalizados = EnumSiNo.No
        TodosPCFinalizados = m_eTodosPCFinalizados
        Exit Property
    End If
    For Each m_ID In Me.ColPCs
        Set m_PC = Me.ColPCs(m_ID)
        m_PCEstado = m_PC.ESTADOCalculado
        If m_PCEstado <> EnumPlanEstado.Finalizado Then
            m_eTodosPCFinalizados = EnumSiNo.No
            TodosPCFinalizados = m_eTodosPCFinalizados
            Exit Property
        End If
        Set m_PC = Nothing
    Next
    m_eTodosPCFinalizados = EnumSiNo.Sí
    TodosPCFinalizados = m_eTodosPCFinalizados
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.TodosPCFinalizados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get AlgunPMDefinido() As EnumSiNo
    
    Dim m_ID As Variant
    Dim m_PM As PM
    Dim m_PMEstado As EnumPlanEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.TienePMs = EnumSiNo.No Then
        m_eAlgunPMDefinido = EnumSiNo.No
        AlgunPMDefinido = m_eAlgunPMDefinido
        Exit Property
    End If
    For Each m_ID In Me.ColPMs
        Set m_PM = Me.ColPMs(m_ID)
        m_PMEstado = m_PM.ESTADOCalculado
        If m_PMEstado = EnumPlanEstado.Definido Then
            m_eAlgunPMDefinido = EnumSiNo.Sí
            AlgunPMDefinido = m_eAlgunPMDefinido
            Exit Property
        End If
        Set m_PM = Nothing
    Next
    m_eAlgunPMDefinido = EnumSiNo.No
    AlgunPMDefinido = m_eAlgunPMDefinido
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.AlgunPMDefinido ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get RequierePlanContingenciaCalculado() As EnumSiNo
    
    Dim m_Estado As EnumRiesgoEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_Estado = Me.EstadoEnum
    If m_Estado = EnumRiesgoEstado.Retirado Then
        m_sRequierePlanContingenciaCalculado = EnumSiNo.No
        RequierePlanContingenciaCalculado = m_sRequierePlanContingenciaCalculado
        Exit Property
    End If
    If m_Estado = EnumRiesgoEstado.Materializado Then
        m_sRequierePlanContingenciaCalculado = EnumSiNo.Sí
        RequierePlanContingenciaCalculado = m_sRequierePlanContingenciaCalculado
        Exit Property
    End If
    If m_Estado = EnumRiesgoEstado.Aceptado Then
        m_sRequierePlanContingenciaCalculado = EnumSiNo.No
        RequierePlanContingenciaCalculado = m_sRequierePlanContingenciaCalculado
        Exit Property
    End If
    If Me.RiesgoAltoOMuyAlto = EnumSiNo.Sí Then
        m_sRequierePlanContingenciaCalculado = EnumSiNo.Sí
        RequierePlanContingenciaCalculado = m_sRequierePlanContingenciaCalculado
        Exit Property
    End If
    m_sRequierePlanContingenciaCalculado = EnumSiNo.No
    RequierePlanContingenciaCalculado = m_sRequierePlanContingenciaCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RequierePlanContingenciaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Let RequierePlanContingenciaCalculado(ByVal m_eVal As EnumSiNo)

    m_sRequierePlanContingenciaCalculado = m_eVal

End Property
Public Property Get RequierePlanContingenciaCalculadoTexto() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.RequierePlanContingenciaCalculado <> 0 Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If Me.RequierePlanContingenciaCalculado = EnumSiNo.Sí Then
            m_sRequierePlanContingenciaCalculadoTexto = "Sí"
        ElseIf Me.RequierePlanContingenciaCalculado = EnumSiNo.No Then
            m_sRequierePlanContingenciaCalculadoTexto = "No"
        End If
        
    End If
    RequierePlanContingenciaCalculadoTexto = m_sRequierePlanContingenciaCalculadoTexto
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RequierePlanContingenciaCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Edicion() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objEdicion Is Nothing Then
        If m_objEdicion.IDEdicion <> Me.IDEdicion Then
            Set m_objEdicion = Nothing
        Else
            Set Edicion = m_objEdicion
            Exit Property
        End If
    End If
    If Me.IDEdicion = "" Then
        m_Error = "No se ha podido determinar la Edición del Riesgo"
        Err.Raise 1000
    End If
    Set m_objEdicion = Constructor.getEdicion(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Edicion = m_objEdicion
     Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.Edicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Let Edicion(valor As Edicion)
    Set m_objEdicion = valor
End Property
Public Property Get ColPMs() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColPMs Is Nothing Then
        Set ColPMs = m_objColPMs
        Exit Property
    End If
    If Me.IDRiesgo = "" Then
        Exit Property
    End If
    Set m_objColPMs = Constructor.getPMs(p_IDRiesgo:=Me.IDRiesgo, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPMs = m_objColPMs
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColPMs ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColPMs(ByVal objNewValue As Variant)

    Set m_objColPMs = objNewValue

End Property

Friend Sub InjectarColPMs(p_colPMs As Scripting.Dictionary)
    Set m_objColPMs = p_colPMs
End Sub

Public Property Get ColPCs() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColPCs Is Nothing Then
        Set ColPCs = m_objColPCs
        Exit Property
    End If
    If Me.IDRiesgo = "" Then
       Exit Property
    End If
    Set m_objColPCs = Constructor.getPCs(Me.IDRiesgo, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColPCs = m_objColPCs
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColPCs ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColPCs(ByVal objNewValue As Variant)

    Set m_objColPCs = objNewValue

End Property

Friend Sub InjectarColPCs(p_colPCs As Scripting.Dictionary)
    Set m_objColPCs = p_colPCs
End Sub




Public Property Get ColAnexos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColAnexos Is Nothing Then
        Set ColAnexos = m_ObjColAnexos
        Exit Property
    End If
    If Me.IDRiesgo = "" Then
        m_Error = "No se conoce el IDRiesgo"
        Err.Raise 1000
    End If
    Set m_ObjColAnexos = Constructor.getAnexosDeRiesgo(Me.IDRiesgo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnexos = m_ObjColAnexos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColAnexos(ByVal sNewValue As Variant)

    Set m_ObjColAnexos = sNewValue

End Property

Public Function ColAnexosTotales(Optional ByRef p_Error As String = "") As Scripting.Dictionary

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Me.IDRiesgo = "" Then
        m_Error = "No se conoce el IDRiesgo"
        err.Raise 1000
    End If
    Set ColAnexosTotales = modFormRiesgoDocumentosHelper.getAnexosDeRiesgoCached(Me.IDRiesgo, p_Error)
    If p_Error <> "" Then
        err.Raise 1000
    End If
    Exit Function

errores:
    If err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColAnexosTotales ha devuelto el error: " & vbNewLine & err.description
    End If
    Me.Error = m_Error
    ' #83: re-raise para que el caller vea Err.Number real. Antes absorbia
    ' el error y la cache global de Err quedaba vacia, rompiendo el contrato
    ' "raise 1000 si IDRiesgo vacio". Con re-raise, el caller's On Error
    ' handler captura Err 1000 correctamente.
    err.Raise err.Number, err.Source, err.description
End Function

' F17: Public Property Set eliminado — ColAnexosTotales ahora es Public Function.
' Antes: Property Get con parámetros (ilegal sintácticamente válido pero conceptualmente
' mal — VBA lo aceptaba pero rompía la simetría con Edicion.ColAnexosTotales que ya
' era Function).

Public Property Get UltimoPM() As PM
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.IDRiesgo = "" Then
        Exit Property
    End If
    Set m_ObjUltimoPM = Constructor.getPMUltimo(Me.IDRiesgo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UltimoPM = m_ObjUltimoPM
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.UltimoPM ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set UltimoPM(ByVal sNewValue As Variant)
    
    Set m_ObjUltimoPM = sNewValue
     
End Property
Public Property Get UltimoPC() As PC

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.IDRiesgo = "" Then
        Exit Property
    End If
    Set m_ObjUltimoPc = Constructor.getPCUltimo(Me.IDRiesgo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UltimoPC = m_ObjUltimoPc
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.UltimoPC ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set UltimoPC(ByVal sNewValue As Variant)
    
    Set m_ObjUltimoPc = sNewValue
     
End Property

Public Property Get ColLineasInformePublicabilidad() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColLineasInformePublicabilidad Is Nothing Then
        Set ColLineasInformePublicabilidad = m_objColLineasInformePublicabilidad
        Exit Property
    End If
    If Me.Edicion.EsActivo = EnumSiNo.No Then
        Exit Property
    End If
    Set m_objColLineasInformePublicabilidad = getColLineasPublicabilidadRiesgo(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColLineasInformePublicabilidad = m_objColLineasInformePublicabilidad
    
     Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColLineasInformePublicabilidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColLineasInformePublicabilidad(ByVal objNewValue As Variant)

    Set m_objColLineasInformePublicabilidad = objNewValue

End Property


Public Property Get HayErrorEnRiesgoCalculado() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Texto As String
    Dim m_ParteColor As String
    Dim dato
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.Edicion.EsActivo = EnumSiNo.No Then
        Exit Property
    End If
    ' NOTE: intentional side effect — clears and re-reads ColLineasInformePublicabilidad
    Set Me.ColLineasInformePublicabilidad = Nothing
    Set m_Col = Me.ColLineasInformePublicabilidad
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Property
    End If
    For Each m_ID In m_Col
        m_Texto = m_Col(m_ID)
        If InStr(1, m_Texto, "||") <> 0 Then
            dato = Split(m_Texto, "||")
            m_ParteColor = dato(1)
            If m_ParteColor = "Rojo" Then
                m_sHayErrorEnRiesgoCalculado = "Sí"
                HayErrorEnRiesgoCalculado = m_sHayErrorEnRiesgoCalculado
                Exit Property
            End If
        End If
    Next
    
    m_sHayErrorEnRiesgoCalculado = "No"
    HayErrorEnRiesgoCalculado = m_sHayErrorEnRiesgoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.HayErrorEnRiesgoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get NombreNodoDescCalculado() As String
    
   
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sNombreNodoDescCalculado = GetNombreParaNodo(Me, EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    NombreNodoDescCalculado = m_sNombreNodoDescCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.NombreNodoDescCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get NombreNodoEstadoCalculado() As String
    
   
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sNombreNodoEstadoCalculado = GetNombreParaNodo(Me, EnumSiNo.No, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    NombreNodoEstadoCalculado = m_sNombreNodoEstadoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.NombreNodoEstadoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLInformeRiesgo() As String
    
    
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set Me.ColLineasInformePublicabilidad = Nothing
    Set m_Col = Me.ColLineasInformePublicabilidad
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLInformeErroresRiesgo = getURLInformeRiesgo(m_Col, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    URLInformeRiesgo = m_sURLInformeErroresRiesgo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.URLInformeRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ErroresRiesgoTexto() As String
    Dim m_HayErrorEnRiesgo As String
    Dim m_Estado As EnumRiesgoEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Me.Error = ""
    m_Error = ""
    
    If Me.EsActivo = EnumSiNo.No Then
        m_sErroresRiesgoTexto = "Edición no activa"
        ErroresRiesgoTexto = m_sErroresRiesgoTexto
        Exit Property
    End If
    m_Estado = Me.EstadoEnum
    If m_Estado = EnumRiesgoEstado.Aceptado Then
        m_sErroresRiesgoTexto = "Riesgo Aceptado. Publicable"
    ElseIf m_Estado = EnumRiesgoEstado.Retirado Then
        m_sErroresRiesgoTexto = "Riesgo Retirado. Publicable"
    Else
        If Me.HayErrorEnRiesgo = "" Then
            Me.HayErrorEnRiesgoGrabar m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            
        End If
        m_HayErrorEnRiesgo = Me.HayErrorEnRiesgo
        
        If m_HayErrorEnRiesgo <> "Sí" Then
            Exit Property
        End If
        m_sErroresRiesgoTexto = getErroresRiesgoTexto(Me.ColLineasInformePublicabilidad, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If m_sErroresRiesgoTexto = "" Then
        m_sErroresRiesgoTexto = "Riesgo publicable"
    End If
    ErroresRiesgoTexto = m_sErroresRiesgoTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ErroresRiesgoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let ErroresRiesgoTexto(ByVal sNewValue As String)

    m_sErroresRiesgoTexto = sNewValue

End Property
Public Property Get Publicable() As EnumSiNo
    
    Dim m_ColMotivosNoPublicable As Scripting.Dictionary
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_bPublicableCalculado Then
        Publicable = m_ePublicable
        Exit Property
    End If
    Set m_ColMotivosNoPublicable = Me.ColMotivosNoPublicable
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColMotivosNoPublicable Is Nothing Then
        m_ePublicable = EnumSiNo.Sí
    Else
        If m_ColMotivosNoPublicable.Count = 0 Then
            m_ePublicable = EnumSiNo.Sí
        Else
            m_ePublicable = EnumSiNo.No
        End If
    End If
    m_bPublicableCalculado = True
    Publicable = m_ePublicable
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.Publicable ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get HaNacidoenEstaEdicion() As EnumSiNo
    
    Dim m_EdicionNace As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eHaNacidoenEstaEdicion <> 0 Then
        HaNacidoenEstaEdicion = m_eHaNacidoenEstaEdicion
        Exit Property
    End If
    Set m_EdicionNace = Me.EdicionEnLaQueNace
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EdicionNace Is Nothing Then
        m_Error = "No se conoce en la edición que nace el riesgo"
        Err.Raise 1000
    End If
    If m_EdicionNace.IDEdicion = Me.IDEdicion Then
        m_eHaNacidoenEstaEdicion = EnumSiNo.Sí
    Else
         m_eHaNacidoenEstaEdicion = EnumSiNo.No
    End If
    
    HaNacidoenEstaEdicion = m_eHaNacidoenEstaEdicion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.HaNacidoenEstaEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get Borrable() As EnumSiNo
    
    Dim m_Form As Form
    Dim m_ProvieneDeRiesgoExterno As EnumSiNo
    Dim m_EdicionPrimeraYActivaConPermisoUsuario As EnumSiNo
    Dim m_HaNacidoEnEstaEdicion As EnumSiNo
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eBorrable <> 0 Then
        Borrable = m_eBorrable
        Exit Property
    End If
    If Me.EsActivo = EnumSiNo.No Then
        m_eBorrable = EnumSiNo.No
        Borrable = m_eBorrable
        Exit Property
    End If
    m_HaNacidoEnEstaEdicion = Me.HaNacidoenEstaEdicion
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_HaNacidoEnEstaEdicion = EnumSiNo.No Then
        m_eBorrable = EnumSiNo.No
        Borrable = m_eBorrable
        Exit Property
    End If
    If Me.UsuarioConectadoAutorizado = EnumSiNo.No Then
        m_eBorrable = EnumSiNo.No
        Borrable = m_eBorrable
        Exit Property
    End If
    m_ProvieneDeRiesgoExterno = Me.ProvieneDeRiesgoExterno
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ProvieneDeRiesgoExterno = EnumSiNo.Sí Then
        If EsTecnico = EnumSiNo.Sí Then
            m_eBorrable = EnumSiNo.No
            Borrable = m_eBorrable
            Exit Property
        
            
        End If
        On Error Resume Next
        Set m_Form = Application.Screen.ActiveForm
        If Err.Number <> 0 Then
            Err.Clear
            m_eBorrable = EnumSiNo.No
            Borrable = m_eBorrable
            Exit Property
        End If
        On Error GoTo errores
        If m_Form.Name <> "FormGestionRiesgos" And m_Form.Name <> "FormRiesgoExternoDetalle" Then
            m_eBorrable = EnumSiNo.No
            Borrable = m_eBorrable
            Exit Property
        End If
            
        
    End If
   
    m_eBorrable = EnumSiNo.Sí
    Borrable = m_eBorrable
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.Borrable ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ProvieneDeRiesgoExterno() As EnumSiNo
    
    Dim m_RiesgoExterno As RiesgoExterno
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eProvieneDeRiesgoExterno <> 0 Then
        ProvieneDeRiesgoExterno = m_eProvieneDeRiesgoExterno
        Exit Property
    End If
    Set m_RiesgoExterno = Me.RiesgoExterno
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RiesgoExterno Is Nothing Then
        m_eProvieneDeRiesgoExterno = EnumSiNo.No
    Else
        m_eProvieneDeRiesgoExterno = EnumSiNo.Sí
    End If
    ProvieneDeRiesgoExterno = m_eProvieneDeRiesgoExterno
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ProvieneDeRiesgoExterno ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get ImpactoGlobalCalculado() As String
    
    Dim m_intMaximo As Integer
    Dim m_intPlazo As Integer
    Dim m_intCalidad As Integer
    Dim m_intCoste As Integer
    Dim m_ImpactoGlobalCalculado As String
    Dim m_intImpactoGlobalCalculado As Integer
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Me.Plazo = Me.Plazo
    If Me.Plazo = "" Or Me.Calidad = "" Or Me.Coste = "" Then
        Exit Property
    End If
    
    If Me.Plazo = "Muy Alto" Then
        m_intPlazo = 5
    ElseIf Me.Plazo = "Alto" Then
        m_intPlazo = 4
    ElseIf Me.Plazo = "Medio" Then
        m_intPlazo = 3
    ElseIf Me.Plazo = "Bajo" Then
        m_intPlazo = 2
    ElseIf Me.Plazo = "Muy Bajo" Then
        m_intPlazo = 1
    End If
    If Me.Calidad = "Muy Alto" Then
        m_intCalidad = 5
    ElseIf Me.Calidad = "Alto" Then
        m_intCalidad = 4
    ElseIf Me.Calidad = "Medio" Then
        m_intCalidad = 3
    ElseIf Me.Calidad = "Bajo" Then
        m_intCalidad = 2
    ElseIf Me.Calidad = "Muy Bajo" Then
        m_intCalidad = 1
    End If
    If Me.Coste = "Muy Alto" Then
        m_intCoste = 5
    ElseIf Me.Coste = "Alto" Then
        m_intCoste = 4
    ElseIf Me.Coste = "Medio" Then
        m_intCoste = 3
    ElseIf Me.Coste = "Bajo" Then
        m_intCoste = 2
    ElseIf Me.Coste = "Muy Bajo" Then
        m_intCoste = 1
    End If
    m_intMaximo = 0
    If m_intPlazo > m_intMaximo Then
        m_intMaximo = m_intPlazo
    End If
    If m_intCalidad > m_intMaximo Then
        m_intMaximo = m_intCalidad
    End If
    If m_intCoste > m_intMaximo Then
        m_intMaximo = m_intCoste
    End If
    If m_intPlazo = m_intCalidad And m_intPlazo = m_intCoste Then
        If m_intPlazo = 5 Or m_intPlazo = 1 Then
            m_intImpactoGlobalCalculado = m_intPlazo
        Else
            m_intImpactoGlobalCalculado = m_intMaximo + 1
        End If
    Else
        m_intImpactoGlobalCalculado = m_intMaximo
    End If
    If m_intImpactoGlobalCalculado = 5 Then
        m_sImpactoGlobalCalculado = "Muy Alto"
    ElseIf m_intImpactoGlobalCalculado = 4 Then
        m_sImpactoGlobalCalculado = "Alto"
    ElseIf m_intImpactoGlobalCalculado = 3 Then
        m_sImpactoGlobalCalculado = "Medio"
    ElseIf m_intImpactoGlobalCalculado = 2 Then
        m_sImpactoGlobalCalculado = "Bajo"
    ElseIf m_intImpactoGlobalCalculado = 1 Then
        m_sImpactoGlobalCalculado = "Muy Bajo"
    End If
    ImpactoGlobalCalculado = m_sImpactoGlobalCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ImpactoGlobalCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ValoracionCalculada() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.ImpactoGlobal = "" Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If Me.Vulnerabilidad = "" Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    m_sValoracionCalculada = Constructor.getValoracion(Me.ImpactoGlobal, Me.Vulnerabilidad, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    ValoracionCalculada = m_sValoracionCalculada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ImpactoGlobal ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get RiesgoAltoOMuyAlto() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.Valoracion = "Alto" Or Me.Valoracion = "Muy Alto" Then
        m_eRiesgoAltoOMuyAlto = EnumSiNo.Sí
    Else
        m_eRiesgoAltoOMuyAlto = EnumSiNo.No
    End If
    RiesgoAltoOMuyAlto = m_eRiesgoAltoOMuyAlto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoAltoOMuyAlto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get RiesgoAltoOMuyAltoTexto() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.RiesgoAltoOMuyAlto = EnumSiNo.Sí Then
        m_sRiesgoAltoOMuyAltoTexto = "Sí"
    ElseIf Me.RiesgoAltoOMuyAlto = EnumSiNo.No Then
        m_sRiesgoAltoOMuyAltoTexto = "No"
    End If
    RiesgoAltoOMuyAltoTexto = m_sRiesgoAltoOMuyAltoTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoAltoOMuyAltoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColMotivosNoPublicable() As Scripting.Dictionary
   
    Dim m_Estado As EnumRiesgoEstado
    Dim m_ObjColMotivosPorRiesgo As Collection
    Dim m_Motivo As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjColMotivosNoPublicable Is Nothing Then
        Set ColMotivosNoPublicable = m_ObjColMotivosNoPublicable
        Exit Property
    End If
    If Me.IDRiesgo = "" Then
        m_Error = "No se puede determinar el IDRiesgo"
        Err.Raise 1000
    End If
    If Me.Edicion Is Nothing Then
        m_Error = "No se conoce la Edición del Riesgo " & Me.CodigoRiesgo
        Err.Raise 1000
    End If
    m_Estado = Me.EstadoEnum
    
    If m_Estado = EnumRiesgoEstado.AceptadoRechazado Then
        m_Motivo = "La mitigación del Riesgo ha sido Aceptar por el Jefe de Proyecto, pero Calidad lo ha rechazado"
        If m_ObjColMotivosPorRiesgo Is Nothing Then
            Set m_ObjColMotivosPorRiesgo = New Collection
        End If
        m_ObjColMotivosPorRiesgo.Add m_Motivo
    End If
    If m_Estado = EnumRiesgoEstado.AceptadoSinVisar Then
        m_Motivo = "La mitigación del Riesgo ha sido Aceptar por el Jefe de Proyecto, pero Calidad aún no lo ha visado"
        If m_ObjColMotivosPorRiesgo Is Nothing Then
            Set m_ObjColMotivosPorRiesgo = New Collection
        End If
        m_ObjColMotivosPorRiesgo.Add m_Motivo
    End If
    
    
    
    If m_Estado = "AceptadoSinVisar" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Motivo = "El Riesgo ha sido Aceptado por el Jefe de Proyecto, pero aún Calidad no lo ha visado"
        If m_ObjColMotivosPorRiesgo Is Nothing Then
            Set m_ObjColMotivosPorRiesgo = New Collection
        End If
        m_ObjColMotivosPorRiesgo.Add m_Motivo
    ElseIf m_Estado = "RetiradoSinVisar" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Motivo = "El Riesgo ha sido Retirado por el Jefe de Proyecto, pero aún Calidad no lo ha visado"
        If m_ObjColMotivosPorRiesgo Is Nothing Then
            Set m_ObjColMotivosPorRiesgo = New Collection
        End If
        m_ObjColMotivosPorRiesgo.Add m_Motivo
    End If
    
    
    If Me.RiesgoAltoOMuyAlto = EnumSiNo.Sí Then
    
    Else
    
    End If
    
    
    If Me.UltimoPM Is Nothing Then
        m_Motivo = "Ha de tener al menos un Plan de Mitigación definido"
        If m_ObjColMotivosPorRiesgo Is Nothing Then
            Set m_ObjColMotivosPorRiesgo = New Collection
        End If
        m_ObjColMotivosPorRiesgo.Add m_Motivo
    End If
    If Me.RiesgoAltoOMuyAlto = EnumSiNo.Sí Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Me.UltimoPM.Estado = Me.UltimoPM.ESTADOCalculadoTexto
        If Me.UltimoPM.Estado = "Definido" Then
            m_Motivo = "El Riesgo es alto o muy alto y el el Plan de Mitigación sólo está definido (Ha de estar al menos Planificado)"
        ElseIf Me.UltimoPM.Estado = "SinAcciones" Then
            m_Motivo = "El Riesgo es alto o muy alto y el el Plan de Mitigación está Sin Acciones (Ha de estar al menos Planificado)"
        End If
        If m_ObjColMotivosPorRiesgo Is Nothing Then
            Set m_ObjColMotivosPorRiesgo = New Collection
        End If
        m_ObjColMotivosPorRiesgo.Add m_Motivo
        If Me.UltimoPC Is Nothing Then
            m_Motivo = "El Riesgo es alto o muy alto y no se ha definido un plan de contingencia"
        Else
            Me.UltimoPC.Estado = Me.UltimoPC.ESTADOCalculadoTexto
            If Me.UltimoPC.Estado = "SinAcciones" Then
                m_Motivo = "El Riesgo es alto o muy alto y no hay un plan de contingencia Activo (Sin Acciones)"
            ElseIf Me.UltimoPC.Estado = "Finalizado" Then
                m_Motivo = "El Riesgo es alto o muy alto y no hay un plan de contingencia Activo (Finalizado)"
            End If
        End If
        If m_Motivo <> "" Then
            If m_ObjColMotivosPorRiesgo Is Nothing Then
            Set m_ObjColMotivosPorRiesgo = New Collection
            End If
            m_ObjColMotivosPorRiesgo.Add m_Motivo
        End If
    End If
    If Not m_ObjColMotivosPorRiesgo Is Nothing Then
        If m_ObjColMotivosNoPublicable Is Nothing Then
            Set m_ObjColMotivosNoPublicable = New Scripting.Dictionary
            m_ObjColMotivosNoPublicable.CompareMode = TextCompare
        End If
        Set m_ObjColMotivosNoPublicable(Me.IDRiesgo) = m_ObjColMotivosPorRiesgo
    End If
    
    Set ColMotivosNoPublicable = m_ObjColMotivosNoPublicable
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColMotivosNoPublicable ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TienePriorizacionUltima() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.Edicion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido obtener la Edición del Riesgo"
        Err.Raise 1000
    End If
    If Not IsNumeric(Me.Priorizacion) Then
        m_Error = "No se ha podido obtener la priorización del Riesgo " & Me.CodigoRiesgo
        Err.Raise 1000
    End If
    If CInt(Me.Priorizacion) = Me.Edicion.colRiesgos.Count Then
        m_eTienePriorizacionUltima = EnumSiNo.Sí
    Else
        m_eTienePriorizacionUltima = EnumSiNo.No
    End If
    TienePriorizacionUltima = m_eTienePriorizacionUltima
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.TienePriorizacionUltima ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TienePriorizacionPrimera() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If CInt(Me.Priorizacion) = 1 Then
        m_eTienePriorizacionPrimera = EnumSiNo.Sí
    Else
        m_eTienePriorizacionPrimera = EnumSiNo.No
    End If
    TienePriorizacionPrimera = m_eTienePriorizacionPrimera
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.TienePriorizacionPrimera ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get RiesgoEdicionAnterior() As riesgo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objRiesgoEdicionAnterior Is Nothing Then
        Set RiesgoEdicionAnterior = m_objRiesgoEdicionAnterior
        Exit Property
    End If
    If Me.Edicion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido obtener la Edición del Riesgo"
        Err.Raise 1000
    End If
    If Me.Edicion.EsPrimeraEdicion = EnumSiNo.Sí Then
        Exit Property
    End If
    Set m_objRiesgoEdicionAnterior = Constructor.getRiesgo(, Me.Edicion.EdicionAnterior.IDEdicion, Me.CodigoRiesgo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoEdicionAnterior = m_objRiesgoEdicionAnterior
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoEdicionAnterior ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get RiesgoEdicionPrimera() As riesgo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objRiesgoEdicionPrimera Is Nothing Then
        Set RiesgoEdicionPrimera = m_objRiesgoEdicionPrimera
        Exit Property
    End If
    If Me.Edicion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido obtener la Edición del Riesgo"
        Err.Raise 1000
    End If
    If Me.Edicion.EsPrimeraEdicion = EnumSiNo.Sí Then
        Set m_objRiesgoEdicionPrimera = Me
        Set RiesgoEdicionPrimera = m_objRiesgoEdicionPrimera
        Exit Property
    End If
    Set m_objRiesgoEdicionPrimera = Constructor.getRiesgo(, Me.Edicion.EdicionPrimera.IDEdicion, Me.CodigoRiesgo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoEdicionPrimera = m_objRiesgoEdicionPrimera
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoEdicionPrimera ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get RiesgoEdicionSiguiente() As riesgo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objRiesgoEdicionSiguiente Is Nothing Then
        Set RiesgoEdicionSiguiente = m_objRiesgoEdicionSiguiente
        Exit Property
    End If
    If Me.Edicion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido obtener la Edición del Riesgo"
        Err.Raise 1000
    End If
    If Me.Edicion.EsUltimaEdicionCalculado = EnumSiNo.Sí Then
        Exit Property
    End If
    Set m_objRiesgoEdicionSiguiente = Constructor.getRiesgo(, Me.Edicion.EdicionSiguiente.IDEdicion, Me.CodigoRiesgo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoEdicionSiguiente = m_objRiesgoEdicionSiguiente
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoEdicionSiguiente ha devuelto el error: " & vbNewLine & Err.Description
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
        m_Error = "El método Riesgo.UsuarioConectadoAutorizado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get RiesgoDeEdicionActiva() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eRiesgoDeEdicionActiva <> 0 Then
        RiesgoDeEdicionActiva = m_eRiesgoDeEdicionActiva
        Exit Property
    End If
    m_eRiesgoDeEdicionActiva = Me.Edicion.EsActivo
    m_Error = Me.Edicion.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RiesgoDeEdicionActiva = m_eRiesgoDeEdicionActiva
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoDeEdicionActiva ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get CodigoRiesgoCalculado() As String
    
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim intOrdinal As Integer
    Dim intOrdinalMax As Integer
    Dim m_CodRiesgo As String
    Dim dato
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Edicion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se sabe a qué Edición pertenece El Riesgo"
        Err.Raise 1000
    End If
    For Each m_IdRiesgo In Me.Edicion.colRiesgos.keys
            Set m_ObjRiesgo = Me.Edicion.colRiesgos(m_IdRiesgo)
            m_CodRiesgo = m_ObjRiesgo.CodigoRiesgo
            If InStr(1, m_CodRiesgo, "R") <> 0 Then
                dato = Split(m_CodRiesgo, "R")
                If IsNumeric(dato(1)) Then
                    intOrdinal = CInt(dato(1))
                    If intOrdinal > intOrdinalMax Then
                        intOrdinalMax = intOrdinal
                    End If
                End If
            End If
            Set m_ObjRiesgo = Nothing
        Next
    m_sCodigoRiesgoCalculado = "R" & Format(intOrdinalMax + 1, "000")
    CodigoRiesgoCalculado = m_sCodigoRiesgoCalculado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.CodigoRiesgoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get CodigoUnicoCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.CodigoRiesgo = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido obtener el código nuevo"
        Err.Raise 1000
    End If
    If Me.Edicion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se puede obtener la edición del Riesgo"
        Err.Raise 1000
    End If
    m_sCodigoUnicoCalculado = Format(Me.Edicion.IDProyecto, "000") & Me.CodigoRiesgo
    CodigoUnicoCalculado = m_sCodigoUnicoCalculado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.CodigoUnicoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get PriorizacionCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    getDatosNuevos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    PriorizacionCalculado = m_sPriorizacionCalculado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.PriorizacionCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaDetectadoCalculado() As String
    
    Dim m_IDRiesgos As Variant
    Dim m_ObjRiesgo As riesgo
        
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.EdicionEnLaQueNace Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido determinar la Edición en la que nace El Riesgo"
        Err.Raise 1000
    End If
    For Each m_IDRiesgos In Me.EdicionEnLaQueNace.colRiesgos.keys
        Set m_ObjRiesgo = Me.EdicionEnLaQueNace.colRiesgos(m_IDRiesgos)
        If m_ObjRiesgo.CodigoUnico = Me.CodigoUnico Then
            m_sFechaDetectadoCalculado = m_ObjRiesgo.FechaDetectado
            Exit For
        End If
        Set m_ObjRiesgo = Nothing
    Next
    FechaDetectadoCalculado = m_sFechaDetectadoCalculado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.FechaDetectadoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaMaterializadoCalculado() As String
    
    Dim m_IDEdicion As Variant
    Dim m_objEdicion As Edicion
    Dim m_IDRiesgos As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_FechaMaterializadoMenor As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    For Each m_IDEdicion In Me.Edicion.Proyecto.colEdiciones.keys
        Set m_objEdicion = Me.Edicion.Proyecto.colEdiciones(m_IDEdicion)
        For Each m_IDRiesgos In m_objEdicion.colRiesgos.keys
            Set m_ObjRiesgo = m_objEdicion.colRiesgos(m_IDRiesgos)
            If m_ObjRiesgo.CodigoUnico = Me.CodigoUnico Then
                If IsDate(m_ObjRiesgo.FechaMaterializado) Then
                    If IsDate(m_FechaMaterializadoMenor) Then
                        If CDate(m_ObjRiesgo.FechaMaterializado) < CDate(m_FechaMaterializadoMenor) Then
                            m_FechaMaterializadoMenor = m_ObjRiesgo.FechaMaterializado
                        End If
                    Else
                        m_FechaMaterializadoMenor = m_ObjRiesgo.FechaMaterializado
                    End If
                End If
            End If
            Set m_ObjRiesgo = Nothing
        Next
        Set m_objEdicion = Nothing
    Next
    FechaMaterializadoCalculado = m_sFechaMaterializadoCalculado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.FechaMaterializadoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EdicionEnLaQueNace() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEdicionEnLaQueNace Is Nothing Then
        Set EdicionEnLaQueNace = m_ObjEdicionEnLaQueNace
        Exit Property
    End If
    Set m_ObjEdicionEnLaQueNace = Constructor.getEdicionEnLaQueNaceElRiesgo(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EdicionEnLaQueNace = m_ObjEdicionEnLaQueNace
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.EdicionEnLaQueNace ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EdicionUltimaEnLaQueExiste() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEdicionUltimaEnLaQueExiste Is Nothing Then
        Set EdicionUltimaEnLaQueExiste = m_ObjEdicionUltimaEnLaQueExiste
        Exit Property
    End If
    Set m_ObjEdicionUltimaEnLaQueExiste = Constructor.getEdicionUltimaEnLaQueExiste(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EdicionUltimaEnLaQueExiste = m_ObjEdicionUltimaEnLaQueExiste
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.EdicionUltimaEnLaQueExiste ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get RiesgoBorradoPorError() As EnumSiNo
    
    Dim m_EdicionEnLaQueNace As Edicion
    Dim m_EdicionUltima As Edicion
    Dim i As Long
    Dim intEdicionEnLaQueNace As Integer
    Dim intEdicionUltima As Integer
    Dim m_Riesgo As riesgo
    
    Dim m_Edicion As Edicion
    Dim m_IDProyecto As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eRiesgoBorradoPorError <> 0 Then
        RiesgoBorradoPorError = m_eRiesgoBorradoPorError
        Exit Property
    End If
    Set m_EdicionEnLaQueNace = Me.EdicionEnLaQueNace
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EdicionEnLaQueNace Is Nothing Then
        m_Error = "No se ha podido detectar la edición en la que nace"
        Err.Raise 1000
    End If
    intEdicionEnLaQueNace = CInt(m_EdicionEnLaQueNace.Edicion)
    Set m_EdicionUltima = Me.Edicion.Proyecto.EdicionUltima
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
     If m_EdicionUltima Is Nothing Then
        m_Error = "No se ha podido detectar la edición ultima"
        Err.Raise 1000
    End If
    m_IDProyecto = Me.Edicion.IDProyecto
    intEdicionUltima = CInt(m_EdicionUltima.Edicion)
    If intEdicionEnLaQueNace = intEdicionUltima Then
        m_eRiesgoBorradoPorError = EnumSiNo.No
        RiesgoBorradoPorError = m_eRiesgoBorradoPorError
        Exit Property
    End If
    
    
    For i = intEdicionEnLaQueNace + 1 To intEdicionUltima
        Set m_Edicion = Constructor.getEdicionV2(m_IDProyecto, CStr(i), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_Riesgo = Constructor.getRiesgo(, m_Edicion.IDEdicion, Me.CodigoRiesgo, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Riesgo Is Nothing Then
            m_eRiesgoBorradoPorError = EnumSiNo.Sí
            RiesgoBorradoPorError = m_eRiesgoBorradoPorError
            Exit Property
        End If
    Next
    m_eRiesgoBorradoPorError = EnumSiNo.No
    
    RiesgoBorradoPorError = m_eRiesgoBorradoPorError
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoBorradoPorError ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get EstadoEnum() As EnumRiesgoEstado
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.Estado = "" Then
        m_eEstadoEnum = Me.ESTADOCalculado
        EstadoEnum = m_eEstadoEnum
        Exit Property
    End If
    m_eEstadoEnum = getEstadoRiesgo(Me.Estado, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    EstadoEnum = m_eEstadoEnum
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.EstadoEnum ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get RiesgoCuandoNace() As riesgo
    
    Dim m_EdicionNace As Edicion
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjRiesgoCuandoNace Is Nothing Then
        Set RiesgoCuandoNace = m_ObjRiesgoCuandoNace
        Exit Property
    End If
    Set m_EdicionNace = Me.EdicionEnLaQueNace
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EdicionNace Is Nothing Then
        Exit Property
    End If
    If m_EdicionNace.IDEdicion = Me.IDEdicion Then
        Set m_ObjRiesgoCuandoNace = Me
        Set RiesgoCuandoNace = m_ObjRiesgoCuandoNace
        Exit Property
    End If
    Set m_ObjRiesgoCuandoNace = Constructor.getRiesgo(, m_EdicionNace.IDEdicion, Me.CodigoRiesgo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoCuandoNace = m_ObjRiesgoCuandoNace
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoCuandoNace ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get RiesgoEnUltimaEdicion() As riesgo
    
    Dim m_EdicionUltima As Edicion
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjRiesgoEnUltimaEdicion Is Nothing Then
        Set RiesgoEnUltimaEdicion = m_ObjRiesgoEnUltimaEdicion
        Exit Property
    End If
    Set m_EdicionUltima = Me.EdicionUltimaEnLaQueExiste
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EdicionUltima Is Nothing Then
        Exit Property
    End If
    If m_EdicionUltima.IDEdicion = Me.IDEdicion Then
        Set m_ObjRiesgoEnUltimaEdicion = Me
        Set RiesgoEnUltimaEdicion = m_ObjRiesgoEnUltimaEdicion
        Exit Property
    End If
    Set m_ObjRiesgoEnUltimaEdicion = Constructor.getRiesgo(, m_EdicionUltima.IDEdicion, Me.CodigoRiesgo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoEnUltimaEdicion = m_ObjRiesgoEnUltimaEdicion
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoEnUltimaEdicion ha devuelto el error: " & vbNewLine & Err.Description
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
    If Me.Edicion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se sabe a qué Edición pertence"
        Err.Raise 1000
    End If
    m_eEsActivo = Me.Edicion.EsActivo
    EsActivo = m_eEsActivo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.EsActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColRiesgosMismoCodigoUnico() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosMismoCodigoUnico Is Nothing Then
        Set ColRiesgosMismoCodigoUnico = m_objColRiesgosMismoCodigoUnico
        Exit Property
    End If
    Set m_objColRiesgosMismoCodigoUnico = Constructor.getRiesgosByCodUnico(Me.CodigoUnico, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRiesgosMismoCodigoUnico = m_objColRiesgosMismoCodigoUnico
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.EsActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosMismoCodigoUnico(ByVal objNewValue As Variant)

    Set m_objColRiesgosMismoCodigoUnico = objNewValue

End Property

Public Property Get ESTADOCalculadoTexto() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sESTADOCalculadoTexto = getRiesgosEstadoCalculadoTexto(Me.ESTADOCalculado, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    ESTADOCalculadoTexto = m_sESTADOCalculadoTexto
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.EstadoCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ContingenciaCalculada() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.UltimoPC Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_sContingenciaCalculada = "No"
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_sContingenciaCalculada = "Sí"
    End If
    ContingenciaCalculada = m_sContingenciaCalculada
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ContingenciaCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get RiesgoExterno() As RiesgoExterno

    Dim m_objEdicion As Edicion
    Dim m_ObjRiesgoExt As RiesgoExterno
    Dim m_IDRiesgoExt As Variant
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjRiesgoExterno Is Nothing Then
        Set RiesgoExterno = m_ObjRiesgoExterno
        Exit Property
    End If
    Set m_objEdicion = Me.Edicion
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objEdicion Is Nothing Then
        m_Error = "No se sabe a qué edición pertence"
        Err.Raise 1000
    End If
    If m_objEdicion.ColRiesgosExternos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_IDRiesgoExt In m_objEdicion.ColRiesgosExternos.keys
        Set m_ObjRiesgoExt = m_objEdicion.ColRiesgosExternos(m_IDRiesgoExt)
        If m_ObjRiesgoExt.IDRiesgo = Me.IDRiesgo Then
            Set m_ObjRiesgoExterno = m_ObjRiesgoExt
            Set RiesgoExterno = m_ObjRiesgoExterno
            Exit Property
        End If
        Set m_ObjRiesgoExt = Nothing
    Next
    Set RiesgoExterno = m_ObjRiesgoExterno
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoExterno ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set RiesgoExterno(ByVal objNewValue As RiesgoExterno)

    Set m_ObjRiesgoExterno = objNewValue

End Property

Public Property Get TienePMs() As EnumSiNo
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Me.ColPMs Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTienePMs = EnumSiNo.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTienePMs = EnumSiNo.Sí
    End If
    TienePMs = m_eTienePMs
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.TienePMs ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TienePCs() As EnumSiNo
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Me.ColPCs Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTienePCs = EnumSiNo.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTienePCs = EnumSiNo.Sí
    End If
     TienePCs = m_eTienePCs
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.TienePCs ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property




Public Property Get ResponsableTelefonica() As ExpedienteResponsable
    
    Dim m_ID As Variant
    Dim m_ObjResponsable As ExpedienteResponsable
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjResponsableTelefonica Is Nothing Then
        Set ResponsableTelefonica = m_ObjResponsableTelefonica
        Exit Property
    End If
    If Me.Edicion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If Me.Edicion.Proyecto Is Nothing Then
        m_Error = Me.Edicion.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If Me.Edicion.Proyecto.ColAutorizados Is Nothing Then
        m_Error = Me.Edicion.Proyecto.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_ID In Me.Edicion.Proyecto.ColAutorizados
        Set m_ObjResponsableTelefonica = Me.Edicion.Proyecto.ColAutorizados(m_ID)
        Exit For
    Next
    Set ResponsableTelefonica = m_ObjResponsableTelefonica
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ResponsableTelefonica ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get RiesgoEnEdicionActiva() As riesgo

    'puede ser un diccionario de RiesgosNC o un RiesgoNC
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objRiesgoEnEdicionActiva Is Nothing Then
        Set RiesgoEnEdicionActiva = m_objRiesgoEnEdicionActiva
        Exit Property
    End If
    
    Set m_objRiesgoEnEdicionActiva = Constructor.getRiesgoEdicionActiva(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoEnEdicionActiva = m_objRiesgoEnEdicionActiva
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoEnEdicionActiva ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set RiesgoEnEdicionActiva(ByVal objNewValue As riesgo)

    Set m_objRiesgoEnEdicionActiva = objNewValue

End Property


Public Property Get FichaRiesgoCompleta() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    With Me
        If Me.Descripcion = "" Then
            m_eFichaRiesgoCompleta = EnumSiNo.No
            FichaRiesgoCompleta = m_eFichaRiesgoCompleta
            Exit Property
        End If
        If Me.CodigoRiesgo = "" Then
            m_eFichaRiesgoCompleta = EnumSiNo.No
            FichaRiesgoCompleta = m_eFichaRiesgoCompleta
            Exit Property
        End If
        If Me.EntidadDetecta = "" Then
            m_eFichaRiesgoCompleta = EnumSiNo.No
            FichaRiesgoCompleta = m_eFichaRiesgoCompleta
            Exit Property
        End If
        If Me.FechaDetectado = "" Then
            m_eFichaRiesgoCompleta = EnumSiNo.No
            FichaRiesgoCompleta = m_eFichaRiesgoCompleta
            Exit Property
        End If
        If Me.Plazo = "" Then
            m_eFichaRiesgoCompleta = EnumSiNo.No
            FichaRiesgoCompleta = m_eFichaRiesgoCompleta
            Exit Property
        End If
        If Me.Coste = "" Then
            m_eFichaRiesgoCompleta = EnumSiNo.No
            FichaRiesgoCompleta = m_eFichaRiesgoCompleta
            Exit Property
        End If
        If Me.Calidad = "" Then
            m_eFichaRiesgoCompleta = EnumSiNo.No
            FichaRiesgoCompleta = m_eFichaRiesgoCompleta
            Exit Property
        End If
        If Me.Vulnerabilidad = "" Then
            m_eFichaRiesgoCompleta = EnumSiNo.No
            FichaRiesgoCompleta = m_eFichaRiesgoCompleta
            Exit Property
        End If
        If Me.Mitigacion = "" Then
            m_eFichaRiesgoCompleta = EnumSiNo.No
            FichaRiesgoCompleta = m_eFichaRiesgoCompleta
            Exit Property
        End If
    End With
    m_eFichaRiesgoCompleta = EnumSiNo.Sí
    FichaRiesgoCompleta = m_eFichaRiesgoCompleta
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.FichaRiesgoCompleta ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasRespuestaCalidadAceptacion() As String
    
    Dim m_FechaCalidad As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not IsDate(Me.FechaJustificacionAceptacionRiesgo) Then
        Exit Property
    End If
    If Not IsDate(Me.FechaAprobacionAceptacionPorCalidad) And Not IsDate(Me.FechaRechazoAceptacionPorCalidad) Then
        Exit Property
    End If
    If IsDate(Me.FechaAprobacionAceptacionPorCalidad) Then
        m_FechaCalidad = Me.FechaAprobacionAceptacionPorCalidad
    Else
        m_FechaCalidad = Me.FechaRechazoAceptacionPorCalidad
    End If
    m_sDiasRespuestaCalidadAceptacion = DateDiff("d", CDate(Me.FechaJustificacionAceptacionRiesgo), CDate(m_FechaCalidad))
    If IsNumeric(m_sDiasRespuestaCalidadAceptacion) Then
        If m_sDiasRespuestaCalidadAceptacion <> "0" Then
            If CDbl(m_sDiasRespuestaCalidadAceptacion) < 1 Then
                m_sDiasRespuestaCalidadAceptacion = CStr(CDbl(m_sDiasRespuestaCalidadAceptacion) * -1)
            End If
            'Stop
       
            
        End If
        
    End If
    DiasRespuestaCalidadAceptacion = m_sDiasRespuestaCalidadAceptacion
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.DiasRespuestaCalidadAceptacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasRespuestaCalidadRetiro() As String
    
    Dim m_FechaCalidad As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not IsDate(Me.FechaJustificacionRetiroRiesgo) Then
        Exit Property
    End If
    If Not IsDate(Me.FechaAprobacionRetiroPorCalidad) And Not IsDate(Me.FechaRechazoRetiroPorCalidad) Then
        Exit Property
    End If
    If IsDate(Me.FechaAprobacionRetiroPorCalidad) Then
        m_FechaCalidad = Me.FechaAprobacionRetiroPorCalidad
    Else
        m_FechaCalidad = Me.FechaRechazoRetiroPorCalidad
    End If
    m_sDiasRespuestaCalidadRetiro = DateDiff("d", CDate(Me.FechaJustificacionRetiroRiesgo), CDate(m_FechaCalidad))
    If IsNumeric(m_sDiasRespuestaCalidadRetiro) Then
        If m_sDiasRespuestaCalidadRetiro <> "0" Then
            If CDbl(m_sDiasRespuestaCalidadRetiro) < 1 Then
                m_sDiasRespuestaCalidadRetiro = CStr(CDbl(m_sDiasRespuestaCalidadRetiro) * -1)
            End If
            'Stop
        
            
        End If
        
    End If
    
    DiasRespuestaCalidadRetiro = m_sDiasRespuestaCalidadRetiro
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.DiasRespuestaCalidadRetiro ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasRespuestaCalidadRetipificacion() As String
    
    Dim m_FechaCalidad As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not IsDate(Me.FechaRiesgoParaRetipificar) Then
        Exit Property
    End If
    If Not IsDate(Me.FechaRiesgoRetipificado) Then
        Exit Property
    End If
    m_FechaCalidad = Me.FechaRiesgoRetipificado
    m_sDiasRespuestaCalidadRetipificacion = DateDiff("d", CDate(Me.FechaRiesgoParaRetipificar), CDate(m_FechaCalidad))
    If IsNumeric(m_sDiasRespuestaCalidadRetipificacion) Then
        If CDbl(m_sDiasRespuestaCalidadRetipificacion) < 1 Then
            m_sDiasRespuestaCalidadRetipificacion = CStr(CDbl(m_sDiasRespuestaCalidadRetipificacion) * -1)
        End If
    End If
    DiasRespuestaCalidadRetipificacion = m_sDiasRespuestaCalidadRetipificacion
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.DiasRespuestaCalidadRetipificacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasSinRespuestaCalidadAceptacionCalculado() As String
    
    Dim m_Fecha As String
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    If Not IsDate(Me.FechaJustificacionAceptacionRiesgo) Then
        Exit Property
    End If
    If IsDate(Me.FechaAprobacionAceptacionPorCalidad) And IsDate(Me.FechaRechazoAceptacionPorCalidad) Then
        Exit Property
    End If
    m_Fecha = CStr(m_FechaRef)
    m_sDiasSinRespuestaCalidadAceptacionCalculado = DateDiff("d", CDate(FechaJustificacionAceptacionRiesgo), CDate(m_Fecha))
    DiasSinRespuestaCalidadAceptacionCalculado = m_sDiasSinRespuestaCalidadAceptacionCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.DiasSinRespuestaCalidadAceptacionCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasSinRespuestaCalidadRetiroCalculado() As String
    
    Dim m_Fecha As String
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not IsDate(Me.FechaJustificacionRetiroRiesgo) Then
        Exit Property
    End If
    If IsDate(Me.FechaAprobacionRetiroPorCalidad) And IsDate(Me.FechaRechazoRetiroPorCalidad) Then
        Exit Property
    End If
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    m_Fecha = CStr(m_FechaRef)
    m_sDiasSinRespuestaCalidadRetiroCalculado = DateDiff("d", CDate(FechaJustificacionRetiroRiesgo), CDate(m_Fecha))
    DiasSinRespuestaCalidadRetiroCalculado = m_sDiasSinRespuestaCalidadRetiroCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.DiasSinRespuestaCalidadRetiroCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasSinRespuestaCalidadRetipificacionCalculado() As String
    
    Dim m_Fecha As String
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not IsDate(Me.FechaRiesgoParaRetipificar) Then
        Exit Property
    End If
    If IsDate(Me.FechaRiesgoRetipificado) Then
        Exit Property
    End If
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    m_Fecha = CStr(m_FechaRef)
    m_sDiasSinRespuestaCalidadRetipificacionCalculado = DateDiff("d", CDate(FechaRiesgoParaRetipificar), CDate(m_Fecha))
    DiasSinRespuestaCalidadRetipificacionCalculado = m_sDiasSinRespuestaCalidadRetipificacionCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.DiasSinRespuestaCalidadRetipificacionCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HTMLRiesgoPorRetipificar() As String
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sHTMLRiesgoPorRetipificar <> "" Then
        HTMLRiesgoPorRetipificar = m_sHTMLRiesgoPorRetipificar
        Exit Property
    End If
    
    m_sHTMLRiesgoPorRetipificar = getHTMLRiesgoPorRetipificar(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    HTMLRiesgoPorRetipificar = m_sHTMLRiesgoPorRetipificar
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.HTMLRiesgoPorRetipificar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HTMLRiesgoMaterializado() As String
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sHTMLRiesgoMaterializado <> "" Then
        HTMLRiesgoMaterializado = m_sHTMLRiesgoMaterializado
        Exit Property
    End If
    
    m_sHTMLRiesgoMaterializado = getHTMLRiesgoMaterializado(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    HTMLRiesgoMaterializado = m_sHTMLRiesgoMaterializado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.HTMLRiesgoMaterializado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get RiesgoAceptadoEnNacimiento() As riesgo
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjRiesgoAceptadoEnNacimiento Is Nothing Then
        Set RiesgoAceptadoEnNacimiento = m_ObjRiesgoAceptadoEnNacimiento
        Exit Property
    End If
    Set m_ObjRiesgoAceptadoEnNacimiento = Constructor.getRiesgoAceptadoEnNacimiento(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoAceptadoEnNacimiento = m_ObjRiesgoAceptadoEnNacimiento
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoAceptadoEnNacimiento ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get RiesgoRetiradoEnNacimiento() As riesgo
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjRiesgoRetiradoEnNacimiento Is Nothing Then
        Set RiesgoRetiradoEnNacimiento = m_ObjRiesgoRetiradoEnNacimiento
        Exit Property
    End If
    Set m_ObjRiesgoRetiradoEnNacimiento = Constructor.getRiesgoRetiradoEnNacimiento(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoRetiradoEnNacimiento = m_ObjRiesgoRetiradoEnNacimiento
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoRetiradoEnNacimiento ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get RiesgoDeBiblioteca() As RiesgoBiblioteca
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Not m_objRiesgoDeBiblioteca Is Nothing Then
        Set RiesgoDeBiblioteca = m_objRiesgoDeBiblioteca
        Exit Property
    End If
    If Me.CodRiesgoBiblioteca = "" Then
        Exit Property
    End If
    Set m_objRiesgoDeBiblioteca = Constructor.getRiesgoBiblioteca(, Me.CodRiesgoBiblioteca, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoDeBiblioteca = m_objRiesgoDeBiblioteca
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoDeBiblioteca ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set RiesgoDeBiblioteca(ByVal objNewValue As RiesgoBiblioteca)

    Set m_objRiesgoDeBiblioteca = objNewValue

End Property


Public Property Get RiesgoMaterializadoUltimo() As RiesgoMaterializacion
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjRiesgoMaterializadoUltimo Is Nothing Then
        Set RiesgoMaterializadoUltimo = m_ObjRiesgoMaterializadoUltimo
        Exit Property
    End If
    
    Set m_ObjRiesgoMaterializadoUltimo = Constructor.getRiesgoMaterializadoUltimo(, , Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoMaterializadoUltimo = m_ObjRiesgoMaterializadoUltimo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoMaterializadoUltimo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get ColMaterializaciones() As Scripting.Dictionary
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Not m_objColMaterializaciones Is Nothing Then
        Set ColMaterializaciones = m_objColMaterializaciones
        Exit Property
    End If
    If Me.IDRiesgo = "" Then
        Exit Property
    End If
    Set m_objColMaterializaciones = Constructor.getRiesgoMaterializadoHistoria(, , Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ColMaterializaciones = m_objColMaterializaciones
    
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColMaterializaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set ColMaterializaciones(p_Valor As Variant)
    Set m_objColMaterializaciones = p_Valor
End Property

Public Property Get colAccionesMitigacionSinCerrar() As Scripting.Dictionary
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Not m_objcolAccionesMitigacionSinCerrar Is Nothing Then
        Set colAccionesMitigacionSinCerrar = m_objcolAccionesMitigacionSinCerrar
        Exit Property
    End If
    If Me.IDRiesgo = "" Then
        Exit Property
    End If
    Set m_objcolAccionesMitigacionSinCerrar = Constructor.getPMAccionesSinCerrar(p_IDRiesgo:=Me.IDRiesgo, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set colAccionesMitigacionSinCerrar = m_objcolAccionesMitigacionSinCerrar
    
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.colAccionesMitigacionSinCerrar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set colAccionesMitigacionSinCerrar(p_Valor As Variant)
    Set m_objcolAccionesMitigacionSinCerrar = p_Valor
End Property

Public Property Get colAccionesContingenciaSinCerrar() As Scripting.Dictionary
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Not m_objcolAccionesContingenciaSinCerrar Is Nothing Then
        Set colAccionesContingenciaSinCerrar = m_objcolAccionesContingenciaSinCerrar
        Exit Property
    End If
    If Me.IDRiesgo = "" Then
        Exit Property
    End If
    Set m_objcolAccionesContingenciaSinCerrar = Constructor.getPCAccionesSinCerrar(p_IDRiesgo:=Me.IDRiesgo, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set colAccionesContingenciaSinCerrar = m_objcolAccionesContingenciaSinCerrar
    
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.colAccionesContingenciaSinCerrar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set colAccionesContingenciaSinCerrar(p_Valor As Variant)
    Set m_objcolAccionesContingenciaSinCerrar = p_Valor
End Property
Public Property Get RiesgoAlgunaVezMaterializado() As EnumSiNo
    
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_eRiesgoAlgunaVezMaterializado <> Empty Then
        RiesgoAlgunaVezMaterializado = m_eRiesgoAlgunaVezMaterializado
        Exit Property
    End If
    Set m_Col = Me.ColMaterializaciones
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_eRiesgoAlgunaVezMaterializado = EnumSiNo.No
    Else
        m_eRiesgoAlgunaVezMaterializado = EnumSiNo.Sí
    End If
    RiesgoAlgunaVezMaterializado = m_eRiesgoAlgunaVezMaterializado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoAlgunaVezMaterializado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get colMaterializacionesPorDecidirNC() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objcolMaterializacionesPorDecidirNC Is Nothing Then
        Set colMaterializacionesPorDecidirNC = m_objcolMaterializacionesPorDecidirNC
        Exit Property
    End If
    Set m_objcolMaterializacionesPorDecidirNC = Constructor.getRiesgoMaterializadoPendientesDecidirNC(, , Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colMaterializacionesPorDecidirNC = m_objcolMaterializacionesPorDecidirNC
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.colMaterializacionesPorDecidirNC ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property





Private Function getHTMLRiesgoPorRetipificar( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim body As String
    Dim errTxt As String
    On Error GoTo errores
    body = GetBodyCorreoRiesgoRequiereRetipificacion(Me, errTxt)
    If errTxt <> "" Then
        Err.Raise 1000
    End If
    getHTMLRiesgoPorRetipificar = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLRiesgoPorRetipificar ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getHTMLRiesgoAceptadoRetirado( _
                                                Optional p_EsAceptado As EnumSiNo = EnumSiNo.Sí, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim body As String
    Dim errTxt As String
    On Error GoTo errores
    If p_EsAceptado = EnumSiNo.Sí Then
        body = GetBodyCorreoRiesgoAceptadoRetirado(Me, EnumSiNo.Sí, errTxt)
    Else
        body = GetBodyCorreoRiesgoAceptadoRetirado(Me, EnumSiNo.No, errTxt)
    End If
    If errTxt <> "" Then
        Err.Raise 1000
    End If
    getHTMLRiesgoAceptadoRetirado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLRiesgoAceptadoRetirado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getHTMLTecnicoRiesgoAceptado( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim body As String
    Dim errTxt As String
    On Error GoTo errores
    body = GetBodyCorreoTecnicoRiesgoAceptado(Me, errTxt)
    If errTxt <> "" Then
        Err.Raise 1000
    End If
    getHTMLTecnicoRiesgoAceptado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLTecnicoRiesgoAceptado ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getHTMLTecnicoRiesgoRetirado( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim body As String
    Dim errTxt As String
    On Error GoTo errores
    body = GetBodyCorreoRiesgoAceptadoRetirado(Me, EnumSiNo.No, errTxt)
    If errTxt <> "" Then
        Err.Raise 1000
    End If
    getHTMLTecnicoRiesgoRetirado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLTecnicoRiesgoRetirado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getHTMLRiesgoRetipificado( _
                                                ByRef p_ObjRiesgoAlInicio As riesgo, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim body As String
    
    On Error GoTo errores
    body = GetBodyCorreoRiesgoRetipificado(Me, p_ObjRiesgoAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getHTMLRiesgoRetipificado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLRiesgoRetipificado ha devuelto el error: " & Err.Description
    End If
End Function



Public Function getHTMLCalidadApruebaRiesgoAceptado( _
                                                        Optional ByRef p_Error As String _
                                                            ) As String
    Dim body As String
    
    On Error GoTo errores
    body = GetBodyCorreoCalidadApruebaRiesgoAceptado(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getHTMLCalidadApruebaRiesgoAceptado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLCalidadApruebaRiesgoAceptado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getHTMLCalidadQuitaAprobacionRiesgoAceptado( _
                                                        Optional ByRef p_Error As String _
                                                            ) As String
    Dim body As String
    
    On Error GoTo errores
    body = GetBodyCorreoCalidadQuitaAprobacionRiesgoAceptado(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getHTMLCalidadQuitaAprobacionRiesgoAceptado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLCalidadQuitaAprobacionRiesgoAceptado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getHTMLCalidadRechazaRiesgoAceptado( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    Dim body As String
    
    On Error GoTo errores
    body = GetBodyCorreoCalidadRechazaRiesgoAceptado(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getHTMLCalidadRechazaRiesgoAceptado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLCalidadRechazaRiesgoAceptado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getHTMLCalidadQuitarRechazoRiesgoAceptado( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    Dim body As String
    
    On Error GoTo errores
    body = GetBodyCorreoCalidadQuitarRechazoRiesgoAceptado(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getHTMLCalidadQuitarRechazoRiesgoAceptado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLCalidadQuitarRechazoRiesgoAceptado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getHTMLCalidadApruebaRiesgoRetirado( _
                                                    Optional ByRef p_Error As String _
                                                            ) As String
    Dim body As String
    
    On Error GoTo errores
    body = GetBodyCorreoCalidadApruebaRiesgoRetirado(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getHTMLCalidadApruebaRiesgoRetirado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLCalidadApruebaRiesgoRetirado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getHTMLCalidadQuitaAprobacionRiesgoRetirado( _
                                                            Optional ByRef p_Error As String _
                                                            ) As String
    Dim body As String
    
    On Error GoTo errores
    body = GetBodyCorreoCalidadQuitaAprobacionRiesgoRetirado(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getHTMLCalidadQuitaAprobacionRiesgoRetirado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLCalidadQuitaAprobacionRiesgoRetirado ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getHTMLCalidadRechazaRiesgoRetirado( _
                                                    Optional ByRef p_Error As String _
                                                            ) As String
    Dim body As String
    
    On Error GoTo errores
    body = GetBodyCorreoCalidadRechazaRiesgoRetirado(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getHTMLCalidadRechazaRiesgoRetirado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLCalidadRechazaRiesgoRetirado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function getHTMLCalidadQuitarRechazoRiesgoRetirado( _
                                                    Optional ByRef p_Error As String _
                                                            ) As String
    Dim body As String
    
    On Error GoTo errores
    body = GetBodyCorreoCalidadQuitarRechazoRiesgoRetirado(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getHTMLCalidadQuitarRechazoRiesgoRetirado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLCalidadQuitarRechazoRiesgoRetirado ha devuelto el error: " & Err.Description
    End If
End Function


Private Function getHTMLRiesgoMaterializado( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim body As String
    
    On Error GoTo errores
    body = GetBodyCorreoRiesgoMaterializado(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getHTMLRiesgoMaterializado = body
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLRiesgoMaterializado ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getDatosNuevos(Optional ByRef p_Error As String) As String
    
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_intOrdinalMaxCodRiesgo As Integer
    Dim m_intOrdinalMaxPriorizacion As Integer
    On Error GoTo errores
    
    Me.Error = ""
    
    If Me.Edicion Is Nothing Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "No se sabe a qué Edición pertenece El Riesgo"
        Err.Raise 1000
    End If
    Set Me.Edicion.colRiesgos = Nothing
    If Not Me.Edicion.colRiesgos Is Nothing Then
        p_Error = Me.Edicion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_IdRiesgo In Me.Edicion.colRiesgos.keys
            Set m_ObjRiesgo = Me.Edicion.colRiesgos(m_IdRiesgo)
            If IsNumeric(Right(m_ObjRiesgo.CodigoRiesgo, 2)) Then
                If CInt(Right(m_ObjRiesgo.CodigoRiesgo, 2)) > m_intOrdinalMaxCodRiesgo Then
                    m_intOrdinalMaxCodRiesgo = CInt(Right(m_ObjRiesgo.CodigoRiesgo, 2))
                End If
            End If
            If IsNumeric(m_ObjRiesgo.Priorizacion) Then
                If CInt(m_ObjRiesgo.Priorizacion) > m_intOrdinalMaxPriorizacion Then
                    m_intOrdinalMaxPriorizacion = CInt(m_ObjRiesgo.Priorizacion)
                End If
            End If
            Set m_ObjRiesgo = Nothing
        Next
    End If
    
    m_sCodigoRiesgoCalculado = "R" & Format(m_intOrdinalMaxCodRiesgo, "00")
    m_sPriorizacionCalculado = CStr(m_intOrdinalMaxPriorizacion + 1)
    If Me.Edicion.Proyecto Is Nothing Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "No se sabe a qué Proyecto pertenece El Riesgo"
        Err.Raise 1000
    End If
    m_sCodigoRiesgoCalculado = Format(Me.Edicion.Proyecto.IDProyecto, "000") & m_sCodigoRiesgoCalculado
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDatosNuevos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function RegistrarDiasAceptacionCalidad(Optional ByRef p_Error As String) As String
    
    Dim m_SQL As String
    Dim m_Dias As String
    
    On Error GoTo errores
    
    Me.Error = ""
    
    m_Dias = Me.DiasRespuestaCalidadAceptacion
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If IsNumeric(m_Dias) Then
        m_SQL = "UPDATE TbRiesgos SET DiasSinRespuestaCalidadAceptacion='" & m_Dias & "' " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    Else
         m_SQL = "UPDATE TbRiesgos SET DiasSinRespuestaCalidadAceptacion=Null " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    End If
    getdb().Execute m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarDiasAceptacionCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function RegistrarDiasAceptacionRetiro(Optional ByRef p_Error As String) As String
    
    Dim m_SQL As String
    Dim m_Dias As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Dias = Me.DiasRespuestaCalidadRetiro
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If IsNumeric(m_Dias) Then
        m_SQL = "UPDATE TbRiesgos SET DiasSinRespuestaCalidadRetiro='" & m_Dias & "' " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    Else
         m_SQL = "UPDATE TbRiesgos SET DiasSinRespuestaCalidadRetiro=Null " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    End If
    getdb().Execute m_SQL
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarDiasAceptacionRetiro ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
    
Public Function RegistrarDiasRetipificacion(Optional ByRef p_Error As String) As String
    
    Dim m_SQL As String
    Dim m_Dias As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Dias = Me.DiasRespuestaCalidadRetipificacion
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If IsNumeric(m_Dias) Then
        m_SQL = "UPDATE TbRiesgos SET DiasSinRespuestaCalidadRetipificacion='" & m_Dias & "' " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    Else
         m_SQL = "UPDATE TbRiesgos SET DiasSinRespuestaCalidadRetipificacion=Null " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    End If
    getdb().Execute m_SQL
        
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarDiasRetipificacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim m_ID As Variant
    Dim m_ObjAnexo As Anexo
    Dim m_SQL As String
    On Error GoTo errores
    
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not Me.ColAnexos Is Nothing Then
        For Each m_ID In Me.ColAnexos
            Set m_ObjAnexo = Me.ColAnexos(m_ID)
            With m_ObjAnexo
                .EliminarAnexo p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End With
            Set m_ObjAnexo = Nothing
        Next
    End If
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "DELETE * FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    getdb().Execute m_SQL

    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Borrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function HayErrorEnRiesgoGrabar( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    Me.HayErrorEnRiesgo = Me.HayErrorEnRiesgoCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "UPDATE TbRiesgos SET HayErrorEnRiesgo = '" & Me.HayErrorEnRiesgo & "' " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.HayErrorEnRiesgoGrabar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function GrabarNombreNodoDesc( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    Me.NombreNodoDesc = Me.NombreNodoDescCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "UPDATE TbRiesgos SET NombreNodoDesc = '" & Me.NombreNodoDesc & "' " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.GrabarNombreNodoDesc ha devuelto el error: " & Err.Description
    End If
End Function

Public Function GrabarNombreNodoEstado( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    Me.NombreNodoEstado = Me.NombreNodoEstadoCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "UPDATE TbRiesgos SET NombreNodoEstado = '" & Me.NombreNodoEstado & "' " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.GrabarNombreNodoEstado ha devuelto el error: " & Err.Description
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
    
    m_SQL = "UPDATE TbRiesgos SET NombreIcono = '" & Me.NombreIcono & "' " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.GrabarNombreIcono ha devuelto el error: " & Err.Description
    End If
End Function
Public Function GrabarColorIcono( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    Me.ColorIcono = Me.ColorIconoCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "UPDATE TbRiesgos SET ColorIcono = '" & Me.ColorIcono & "' " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.GrabarColorIcono ha devuelto el error: " & Err.Description
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
    
    m_SQL = "UPDATE TbRiesgos SET Estado = '" & Me.Estado & "' " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.GrabarEstado ha devuelto el error: " & Err.Description
    End If
End Function
Public Function DescripcionRepetida( _
                                    p_Descripcion As String, _
                                    Optional p_IDRiesgoAExcluir As String, _
                                    Optional ByRef p_Error As String _
                                    ) As EnumSiNo
                               
    Dim m_ID As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_Descripcion As String
    Dim m_objEdicion As Edicion
    
    On Error GoTo errores
    Set m_objEdicion = Me.Edicion
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objEdicion Is Nothing Then
        p_Error = "No se sabe a qué edición pertence este Riesgo"
        Err.Raise 1000
    End If
    If m_objEdicion.colRiesgos Is Nothing Then
        DescripcionRepetida = EnumSiNo.No
        Exit Function
    End If
    For Each m_ID In m_objEdicion.colRiesgos.keys
        If Nz(m_ID, "") = "" Then
            GoTo siguienteRiesgo
        End If
        Set m_ObjRiesgo = m_objEdicion.colRiesgos(m_ID)
        m_Descripcion = m_ObjRiesgo.Descripcion
        If p_Descripcion = m_Descripcion Then
            If p_IDRiesgoAExcluir <> "" Then
                If p_IDRiesgoAExcluir = m_ObjRiesgo.IDRiesgo Then
                    Set m_ObjRiesgo = Nothing
                    GoTo siguienteRiesgo
                End If
            End If
            DescripcionRepetida = EnumSiNo.Sí
            Set m_ObjRiesgo = Nothing
            Exit Function
        End If
        Set m_ObjRiesgo = Nothing
siguienteRiesgo:
    Next
    
    DescripcionRepetida = EnumSiNo.No
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DescripcionRepetida ha devuelto el error: " & Err.Description
    End If
End Function

Public Function DescripcionYCausaRaizRepetidas( _
                                                p_Descripcion As String, _
                                                p_CausaRaiz As String, _
                                                Optional p_IDRiesgoAExcluir As String, _
                                                Optional ByRef p_Error As String _
                                                ) As EnumSiNo
                               
    
    Dim m_objEdicion As Edicion
    
    On Error GoTo errores
    Set m_objEdicion = Me.Edicion
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objEdicion Is Nothing Then
        p_Error = "No se sabe a qué edición pertence este Riesgo"
        Err.Raise 1000
    End If
    DescripcionYCausaRaizRepetidas = m_objEdicion.DescripcionYCausaRaizRepetida(p_Descripcion, p_CausaRaiz, p_IDRiesgoAExcluir, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DescripcionYCausaRaizRepetidas ha devuelto el error: " & Err.Description
    End If
End Function


                                        
Public Function AceptacionRechazo( _
                                        Optional ByRef p_Estado As EnumRiesgoEstado, _
                                        Optional p_Fecha As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "El Riesgo no está en una edición activa"
        Err.Raise 1000
    End If
    If p_Estado = Empty Then
        p_Estado = Me.ESTADOCalculado
    End If
    
    
    If p_Estado <> EnumRiesgoEstado.AceptadoSinVisar Then
        p_Error = "La Aceptación del Riesgo no está pendiente de visado"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    Me.FechaRechazoAceptacionPorCalidad = p_Fecha
    Me.FechaAprobacionAceptacionPorCalidad = ""
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaRechazoAceptacionPorCalidad") = Me.FechaRechazoAceptacionPorCalidad
            .Fields("FechaAprobacionAceptacionPorCalidad") = Null
            .Fields("DiasSinRespuestaCalidadAceptacion") = Me.DiasSinRespuestaCalidadAceptacionCalculado
            .Fields("Estado") = "AceptadoRechazado"
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    p_Estado = EnumRiesgoEstado.AceptadoRechazado
    
    If Me.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
        EnviarCorreoCalidadRechazaRiesgoAceptado Me, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AceptacionRechazo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
                                        
End Function

Public Function AceptacionRechazoQuitar( _
                                        Optional ByRef p_Estado As EnumRiesgoEstado, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "El Riesgo no está en una edición activa"
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        p_Error = "Operación solo autorizada a Calidad"
        Err.Raise 1000
    End If
    If p_Estado = Empty Then
        p_Estado = Me.ESTADOCalculado
    End If
    
    If p_Estado <> EnumRiesgoEstado.AceptadoRechazado Then
        p_Error = "La Aceptación del Riesgo no está Rechazado"
        Err.Raise 1000
    End If
    Me.FechaAprobacionAceptacionPorCalidad = ""
    Me.FechaRechazoAceptacionPorCalidad = ""
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaAprobacionAceptacionPorCalidad") = Null
            .Fields("FechaRechazoAceptacionPorCalidad") = Null
            .Fields("DiasSinRespuestaCalidadAceptacion") = Null
            
            .Fields("Estado") = "AceptadoSinVisar"
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    p_Estado = EnumRiesgoEstado.AceptadoSinVisar
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AceptacionRechazoQuitar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
                                        
End Function





Public Function RetiroRechazo( _
                                Optional p_Estado As EnumRiesgoEstado, _
                                Optional p_Fecha As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "El Riesgo no está en una edición activa"
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        p_Error = "Operación solo autorizada a Calidad"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    If p_Estado = Empty Then
        p_Estado = Me.ESTADOCalculado
    End If
    If p_Estado <> EnumRiesgoEstado.RetiradoSinVisar Then
        p_Error = "El Riesgo no está pendiente de visado"
        Err.Raise 1000
    End If
    Me.FechaAprobacionRetiroPorCalidad = ""
    Me.FechaRechazoRetiroPorCalidad = p_Fecha
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaRechazoRetiroPorCalidad") = Me.FechaRechazoRetiroPorCalidad
            .Fields("FechaAprobacionRetiroPorCalidad") = Null
            .Fields("DiasSinRespuestaCalidadRetiro") = Null
            .Fields("Estado") = "RetiradoRechazado"
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Me.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
        EnviarCorreoCalidadRechazaRiesgoRetirado Me, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RetiroRechazo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
                                        
End Function

Public Function RetiroRechazoQuitar( _
                                        Optional p_Estado As EnumRiesgoEstado, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "El Riesgo no está en una edición activa"
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        p_Error = "Operación solo autorizada a Calidad"
        Err.Raise 1000
    End If
    If p_Estado = Empty Then
        p_Estado = Me.ESTADOCalculado
    End If
    If p_Estado <> EnumRiesgoEstado.RetiradoRechazado Then
        p_Error = "El Retiro del Riesgo no está rechazado por Calidad"
        Err.Raise 1000
    End If
    Me.FechaAprobacionRetiroPorCalidad = ""
    Me.FechaRechazoRetiroPorCalidad = ""
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaAprobacionRetiroPorCalidad") = Null
            .Fields("FechaRechazoRetiroPorCalidad") = Null
            .Fields("DiasSinRespuestaCalidadRetiro") = Null
            .Fields("Estado") = "RetiradoSinVisar"
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Me.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
        EnviarCorreoCalidadQuitarRechazoRiesgoRetirado Me, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RetiroRechazoQuitar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
                                        
End Function



Public Function GrabarPrioridad( _
                                m_Priorizacion As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_SQL As String
    
    On Error GoTo errores
    Me.Priorizacion = m_Priorizacion
    m_SQL = "UPDATE TbRiesgos " & _
            "SET " & _
            "TbRiesgos.Priorizacion =" & Me.Priorizacion & " " & _
            "WHERE (((TbRiesgos.IDRiesgo)=" & Me.IDRiesgo & "));"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getdb().Execute m_SQL
    Me.GrabarEstado
    Me.GrabarNombreNodoDesc
    Me.GrabarNombreNodoEstado
    Me.GrabarColorIcono
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarPrioridad ha devuelto el error: " & Err.Description
    End If
                                                    
End Function

Public Function CambiarPrioridadADosRiesgos( _
                                            ByRef p_ObjRiesgo1 As riesgo, _
                                            ByRef p_ObjRiesgo2 As riesgo, _
                                            Optional ByRef p_Error As String _
                                            ) As String

    Dim m_SQL As String
    
    On Error GoTo errores
    Me.Error = ""
    m_SQL = "UPDATE TbRiesgos " & _
            "SET " & _
            "TbRiesgos.Priorizacion =" & p_ObjRiesgo2.Priorizacion & " " & _
            "WHERE (((TbRiesgos.IDRiesgo)=" & p_ObjRiesgo1.IDRiesgo & "));"
    InvalidarCacheRiesgo p_ObjRiesgo1.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    InvalidarCacheRiesgo p_ObjRiesgo2.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getdb().Execute m_SQL
    
    m_SQL = "UPDATE TbRiesgos " & _
            "SET " & _
            "TbRiesgos.Priorizacion =" & p_ObjRiesgo1.Priorizacion & " " & _
            "WHERE (((TbRiesgos.IDRiesgo)=" & p_ObjRiesgo2.IDRiesgo & "));"
    getdb().Execute m_SQL
    
    Set p_ObjRiesgo1 = getRiesgo(p_ObjRiesgo1.IDRiesgo, , , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set p_ObjRiesgo2 = getRiesgo(p_ObjRiesgo2.IDRiesgo, , , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjEdicionActiva = Constructor.getEdicion(m_ObjEdicionActiva.IDEdicion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CambiarPrioridadADosRiesgos ha devuelto el error: " & Err.Description
    End If

End Function

Public Function EstablecerPriorizacion( _
                                        p_Priorizacion As Single, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    
    Dim m_SQL As String
    
    On Error GoTo errores
    
    
    If Not IsNumeric(p_Priorizacion) Then
        p_Error = "Se ha de indicar p_Priorizacion"
        Err.Raise 1000
    End If
    If p_Priorizacion = 0 Then
        m_SQL = "UPDATE TbRiesgos " & _
            "SET Priorizacion=NULL " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    Else
        m_SQL = "UPDATE TbRiesgos " & _
            "SET Priorizacion=" & p_Priorizacion & " " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    End If
    
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    getdb().Execute m_SQL
    Set m_ObjEdicionActiva = Constructor.getEdicion(Me.Edicion.IDEdicion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerPriorizacion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function MotivoNoOK( _
                                Optional ByRef p_ObjRiesgoAlInicio As riesgo, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    
    Dim m_DescripcionCausaRaizRepetidas As EnumSiNo
    Dim m_Estado As EnumRiesgoEstado
    Dim m_PriorizacionCorrecta As EnumSiNo
    On Error GoTo errores
    With Me
        If .Descripcion = "" Then
            MotivoNoOK = "Falta Descripcion"
            Exit Function
        End If
        If Not IsDate(.FechaDetectado) Then
            MotivoNoOK = "Falta FechaDetectado"
            Exit Function
        End If
        If .Edicion Is Nothing Then
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "No se puede determinar la edición"
            Exit Function
        End If
        If .Edicion.Proyecto Is Nothing Then
            p_Error = Me.Edicion.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "No se puede determinar la gestión de riesgos"
            Exit Function
        End If
        
        If .EsActivo <> EnumSiNo.Sí Then
            MotivoNoOK = "El Riesgo no está activo"
            Exit Function
        End If
        If .UsuarioConectadoAutorizado = EnumSiNo.No Then
            MotivoNoOK = "El perfil de su usuario no tiene autorización para realizar esta operación"
            Exit Function
        End If
        If .DetectadoPor = "" Then
            MotivoNoOK = "Falta DetectadoPor"
            Exit Function
        End If
        If .EntidadDetecta = "" Then
            MotivoNoOK = "Falta EntidadDetecta"
            Exit Function
        End If
        If .Plazo = "" Then
            MotivoNoOK = "Falta Plazo"
            Exit Function
        End If
        If .Coste = "" Then
            MotivoNoOK = "Falta Coste"
            Exit Function
        End If
        If .Calidad = "" Then
            MotivoNoOK = "Falta Calidad"
            Exit Function
        End If
        If .Vulnerabilidad = "" Then
            MotivoNoOK = "Falta Vulnerabilidad"
            Exit Function
        End If
        If .Mitigacion = "" Then
            MotivoNoOK = "Falta Mitigacion"
            Exit Function
        End If
'        If IsNumeric(.Priorizacion) Then
'            m_PriorizacionCorrecta = EsPriorizacionCorrecta(Me.IDEdicion, Me.Priorizacion, Me.IDRiesgo, p_Error)
'            If p_Error <> "" Then
'                Err.Raise 1000
'            End If
'            If m_PriorizacionCorrecta = EnumSiNo.no Then
'                MotivoNoOK = "Priorización no es correcta"
'                Exit Function
'            End If
'
'        Else
'
'            If EsTecnico = EnumSiNo.Sí Then
'                MotivoNoOK = "Priorización ha de ser un número"
'                Exit Function
'            End If
'        End If
        If .Origen = "" Then
            MotivoNoOK = "Falta el origen"
            Exit Function
        End If
        If EsTecnico = EnumSiNo.Sí Then
            If .Origen = "Oferta" And p_ObjRiesgoAlInicio Is Nothing Then
                MotivoNoOK = "Los riesgos que proceden de la Oferta los mete Calidad en la primera edición desde su formulario correspondiente"
                Exit Function
            End If
        End If
        
        
        
        
        If .Origen <> "Oferta" And .Origen <> "UTE" And .Origen <> "Ejecución" And .Origen <> "Suministrador" And .Origen <> "Pedido" Then
            MotivoNoOK = "El Origen ha de ser de uno de estos tipos: " & vbNewLine & _
                        "-Oferta " & vbNewLine & _
                        "-UTE " & vbNewLine & _
                        "-Ejecución " & vbNewLine & _
                        "-Suministrador " & vbNewLine & _
                        "-Pedido "
            Exit Function
        End If
        If .CodigoRiesgo = "" Then
            If Not p_ObjRiesgoAlInicio Is Nothing Then
                MotivoNoOK = "Falta CodigoRiesgo"
                Exit Function
            Else
                .CodigoRiesgo = .CodigoRiesgoCalculado
                p_Error = .Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If .CodigoRiesgo = "" Then
                    MotivoNoOK = "Falta CodigoRiesgo"
                    Exit Function
                End If
            End If
            
        End If
        If .CodigoUnico = "" Then
            If Not p_ObjRiesgoAlInicio Is Nothing Then
                MotivoNoOK = "Falta CodigoUnico"
                Exit Function
            Else
                .CodigoUnico = .CodigoUnicoCalculado
                p_Error = .Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If .CodigoUnico = "" Then
                    MotivoNoOK = "Falta CodigoUnico"
                    Exit Function
                End If
            End If
            
        End If
        
        
        If .Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
            'no puede estar en ningún riesgo de la oferta que pudiera tener esta gestión de riesgos
            
            If p_ObjRiesgoAlInicio Is Nothing Then
                If .Edicion.Proyecto.DescripcionEnRiesgoNoTrasladadoRepetida(.Descripcion, p_Error) = EnumSiNo.Sí Then
                    MotivoNoOK = "La Descripción de este Riesgo corresponde con otro que proviene de la Oferta que no ha sido trasladado"
                    Exit Function
                End If
                If .DescripcionRepetida(.Descripcion, , p_Error) = EnumSiNo.Sí Then
                    MotivoNoOK = "La Descripción de este Riesgo ya está repetida"
                    Exit Function
                End If
            Else
                If .Descripcion <> p_ObjRiesgoAlInicio.Descripcion Then
                    
                    If .Edicion.Proyecto.DescripcionEnRiesgoNoTrasladadoRepetida(.Descripcion, p_Error) = EnumSiNo.Sí Then
                        MotivoNoOK = "La Descripción de este Riesgo corresponde con otro que proviene de la Oferta que no ha sido trasladado"
                        Exit Function
                    End If
                End If
            End If
            
        Else
            If .CausaRaiz = "" Then
                MotivoNoOK = "Falta CausaRaiz"
                Exit Function
            End If
            m_DescripcionCausaRaizRepetidas = .Edicion.DescripcionYCausaRaizRepetidaDeRiesgosNoTrasladados( _
                                            .Descripcion, .CausaRaiz, p_Error)
            If m_DescripcionCausaRaizRepetidas = EnumSiNo.Sí Then
                MotivoNoOK = "La Descripción y la causa raíz de este Riesgo corresponden con otro que proviene de la Oferta que no ha sido trasladado"
                Exit Function
            End If
            m_DescripcionCausaRaizRepetidas = 0
            If p_ObjRiesgoAlInicio Is Nothing Then
                m_DescripcionCausaRaizRepetidas = .DescripcionYCausaRaizRepetidas( _
                                                .Descripcion, .CausaRaiz, , p_Error)
            Else
                If .Descripcion <> p_ObjRiesgoAlInicio.Descripcion Or .CausaRaiz <> p_ObjRiesgoAlInicio.CausaRaiz Then
                    m_DescripcionCausaRaizRepetidas = .DescripcionYCausaRaizRepetidas( _
                                                .Descripcion, .CausaRaiz, , p_Error)
                End If
            End If
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_DescripcionCausaRaizRepetidas = EnumSiNo.Sí Then
                MotivoNoOK = "La Descripción y la causa raíz de este Riesgo ya están repetidas"
                Exit Function
            End If
        End If
        m_Estado = .ESTADOCalculado

        'm_Estado = m_EstadoCalculadoTexto
        'si aparece como retirado y no tiene los visados no puede ser
        If m_Estado = EnumRiesgoEstado.AceptadoSinJustificar Then
            MotivoNoOK = "Se ha marcado como Acpetado y no se ha Justificado"
            Exit Function
        End If
        If m_Estado = EnumRiesgoEstado.AceptadoSinVisar Then
            MotivoNoOK = "Se ha marcado como Acpetado y no se ha visado por Calidad"
            Exit Function
        End If
        If m_Estado = EnumRiesgoEstado.AceptadoRechazado Then
            MotivoNoOK = "Se ha marcado como Acpetado y se ha Rechazado por Calidad"
            Exit Function
        End If
        
        If m_Estado = EnumRiesgoEstado.RetiradoSinJustificar Then
            MotivoNoOK = "Se ha marcado como Retirado y no se ha Justificado"
            Exit Function
        End If
        If m_Estado = EnumRiesgoEstado.RetiradoSinVisar Then
            MotivoNoOK = "Se ha marcado como Retirado y no se ha visado por Calidad"
            Exit Function
        End If
        If m_Estado = EnumRiesgoEstado.RetiradoRechazado Then
            MotivoNoOK = "Se ha marcado como Retirado y no se ha Rechazado por Calidad"
            Exit Function
        End If
        
        
        
    End With
    
   
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
     End If
    
End Function
Public Function RegistrarRiesgoMaterializado( _
                                            p_Estado As String, _
                                            Optional ByRef p_Error As String, _
                                            Optional ByVal p_IDPlanContingencia As String _
                                            ) As String
                            
    Dim m_RiesgoMaterializado As RiesgoMaterializacion
    Dim m_RiesgoMaterializadoUltimo As RiesgoMaterializacion
    Dim blnGenerar As Boolean
    If Me.IDRiesgo = "" Then
        p_Error = "No se ha dado de alta el riesgo aún"
        Err.Raise 1000
    End If
    If p_Estado = "" Then
        p_Error = "Falta el estado"
        Err.Raise 1000
    End If
    If p_IDPlanContingencia <> "" Then
        If ValidarYResolverPlanContingenciaMaterializacion( _
            Me.IDRiesgo, _
            p_IDPlanContingencia, _
            p_Error) <> EnumSiNo.Sí Then
            Err.Raise 1000
        End If
    End If
    Set m_RiesgoMaterializadoUltimo = Me.RiesgoMaterializadoUltimo
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RiesgoMaterializadoUltimo Is Nothing Then
        If Me.FechaMaterializado = "" Then
            p_Error = "El Riesgo no aparece materializado"
            Err.Raise 1000
        End If
        blnGenerar = True
    Else
        If IsDate(Me.FechaMaterializado) Then
            If m_RiesgoMaterializadoUltimo.EsMaterializacion = "No" Then
                blnGenerar = True
            Else
                blnGenerar = False
            End If
        Else
            If m_RiesgoMaterializadoUltimo.EsMaterializacion = "Sí" Then
                blnGenerar = True
            Else
                blnGenerar = False
            End If
        End If
    End If
    If blnGenerar = False Then
        Exit Function
    End If
    Set m_RiesgoMaterializado = New RiesgoMaterializacion
    With m_RiesgoMaterializado
        .IDProyecto = Me.Edicion.IDProyecto
        .IDEdicion = Me.IDEdicion
        .CodigoRiesgo = Me.CodigoRiesgo
        .Estado = p_Estado
        If IsDate(Me.FechaMaterializado) Then
            .Fecha = Me.FechaMaterializado
            .EsMaterializacion = "Sí"
            .IDPlanContingencia = p_IDPlanContingencia
        Else
            .Fecha = Date
            .EsMaterializacion = "No"
        End If
        .RegistrarAlta p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método RegistrarRiesgoMaterializado ha devuelto el error: " & Err.Description
     End If
End Function
Public Function RegistrarRiesgoMaterializadoQuitar( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
                            
    Dim m_RiesgoMaterializado As RiesgoMaterializacion
    Dim m_RiesgoMaterializadoUltimo As RiesgoMaterializacion
    Dim blnGenerar As Boolean
    If Me.IDRiesgo = "" Then
        p_Error = "No se ha dado de alta el riesgo aún"
        Err.Raise 1000
    End If
    
    Set m_RiesgoMaterializadoUltimo = Me.RiesgoMaterializadoUltimo
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RiesgoMaterializadoUltimo Is Nothing Then
        If Me.FechaMaterializado <> "" Then
            p_Error = "El Riesgo aparece materializado"
            Err.Raise 1000
        End If
        blnGenerar = True
    Else
        If IsDate(Me.FechaMaterializado) Then
            If m_RiesgoMaterializadoUltimo.EsMaterializacion = "No" Then
                blnGenerar = True
            Else
                blnGenerar = False
            End If
        Else
            If m_RiesgoMaterializadoUltimo.EsMaterializacion = "Sí" Then
                blnGenerar = True
            Else
                blnGenerar = False
            End If
        End If
    End If
    If blnGenerar = False Then
        Exit Function
    End If
    Set m_RiesgoMaterializado = New RiesgoMaterializacion
    With m_RiesgoMaterializado
        .IDProyecto = Me.Edicion.IDProyecto
        .IDEdicion = Me.IDEdicion
        .CodigoRiesgo = Me.CodigoRiesgo
        If IsDate(Me.FechaMaterializado) Then
            .Fecha = Me.FechaMaterializado
            .EsMaterializacion = "Sí"
        Else
            .Fecha = Date
            .EsMaterializacion = "No"
        End If
        .RegistrarAlta p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método RegistrarRiesgoMaterializadoQuitar ha devuelto el error: " & Err.Description
     End If
End Function
Public Function Registrar( _
                            Optional p_Comprobando As EnumSiNo = EnumSiNo.Sí, _
                            Optional p_ObjRiesgoAlInicio As riesgo, _
                            Optional ByRef p_Error As String _
                            ) As String
                          
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim EsAceptacion As EnumSiNo
    Dim m_ParaTareaRetipificacion As EnumSiNo
    Dim m_ParaTareaRetipificado As EnumSiNo
    Dim m_FechaRef As Date
    Dim m_RiesgoMaterializado As RiesgoMaterializacion
    Dim m_RequiereRiesgoDeBibliotecaCalculado As EnumSiNo
    
    Dim m_ParaRetipificarAntes As EnumSiNo
    Dim m_ParaRetipificarAhora As EnumSiNo
    On Error GoTo errores
    
    Me.Error = ""
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    
    If p_Comprobando = EnumSiNo.Sí Then
        m_MotivoNoOK = MotivoNoOK(p_ObjRiesgoAlInicio, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = m_MotivoNoOK
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    With Me
        .ImpactoGlobal = .ImpactoGlobalCalculado
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        .Valoracion = .ValoracionCalculada
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        .Contingencia = .ContingenciaCalculada
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
       .RequierePlanContingencia = .RequierePlanContingenciaCalculadoTexto
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If .CodigoRiesgo = "" Then
            .CodigoRiesgo = .CodigoRiesgoCalculado
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If .CodigoUnico = "" Then
            .CodigoUnico = .CodigoUnicoCalculado
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If .FechaEstado = "" Then
            .FechaEstado = CStr(m_FechaRef)
        End If

        If Not p_ObjRiesgoAlInicio Is Nothing Then
            If Not IsDate(.FechaRetirado) Then
                .FechaRechazoRetiroPorCalidad = ""
                .FechaAprobacionRetiroPorCalidad = ""
                .FechaJustificacionRetiroRiesgo = ""
                .JustificacionRetiroRiesgo = ""
            End If
            If .Mitigacion <> "Aceptar" Then
                .FechaMitigacionAceptar = ""
                .FechaAprobacionAceptacionPorCalidad = ""
                .FechaRechazoAceptacionPorCalidad = ""
                .JustificacionAceptacionRiesgo = ""
                .FechaJustificacionAceptacionRiesgo = ""
            End If
        End If
        .Estado = .ESTADOCalculadoTexto
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    If p_ObjRiesgoAlInicio Is Nothing Then
        Me.IDRiesgo = Me.IDRiesgoCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_SQL = "TbRiesgos"
    Else
        If Me.IDRiesgo = "" Then
            p_Error = "Es una edición y no existeIDRiesgo "
            Err.Raise 1000
        End If
        m_SQL = "SELECT * FROM TbRiesgos " & _
                "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
        
        
    End If
    m_RequiereRiesgoDeBibliotecaCalculado = Me.RequiereRiesgoDeBibliotecaCalculado
    If m_RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
        Me.RequiereRiesgoDeBiblioteca = "Sí"
    Else
        Me.RequiereRiesgoDeBiblioteca = "No"
    End If
    InvalidarCacheEdicion Me.IDEdicion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not p_ObjRiesgoAlInicio Is Nothing Then
            If .EOF Then
                p_Error = "No se ha podido obtener los datos"
                Err.Raise 1000
            End If
        End If
        If p_ObjRiesgoAlInicio Is Nothing Then
            .AddNew
                If Me.IDRiesgo <> "" Then
                    .Fields("IDRiesgo") = Me.IDRiesgo
                End If
                If Me.IDEdicion <> "" Then
                    .Fields("IDEdicion") = Me.IDEdicion
                End If
                If Me.CodigoUnico <> "" Then
                    .Fields("CodigoUnico") = Me.CodigoUnico
                End If
                If Me.CodigoRiesgo <> "" Then
                    .Fields("CodigoRiesgo") = Me.CodigoRiesgo
                End If
                If Me.FechaDetectado <> "" Then
                    .Fields("FechaDetectado") = Me.FechaDetectado
                End If
                If Me.DetectadoPor <> "" Then
                    .Fields("DetectadoPor") = Me.DetectadoPor
                End If
                If Me.EntidadDetecta <> "" Then
                    .Fields("EntidadDetecta") = Me.EntidadDetecta
                End If
                If Me.Plazo <> "" Then
                    .Fields("Plazo") = Me.Plazo
                End If
                If Me.Calidad <> "" Then
                    .Fields("Calidad") = Me.Calidad
                End If
                If Me.Coste <> "" Then
                    .Fields("Coste") = Me.Coste
                End If
                If Me.ImpactoGlobal <> "" Then
                    .Fields("ImpactoGlobal") = Me.ImpactoGlobal
                End If
                If Me.Vulnerabilidad <> "" Then
                    .Fields("Vulnerabilidad") = Me.Vulnerabilidad
                End If
                If Me.Valoracion <> "" Then
                    .Fields("Valoracion") = Me.Valoracion
                End If
                If Me.Mitigacion <> "" Then
                    .Fields("Mitigacion") = Me.Mitigacion
                End If
                If Me.Contingencia <> "" Then
                    .Fields("Contingencia") = Me.Contingencia
                End If
                If Me.RequierePlanContingencia <> "" Then
                    .Fields("RequierePlanContingencia") = Me.RequierePlanContingencia
                End If
                If Me.Descripcion <> "" Then
                    .Fields("Descripcion") = Me.Descripcion
                End If
                If Me.CausaRaiz <> "" Then
                    .Fields("CausaRaiz") = Me.CausaRaiz
                End If
                If Me.Estado <> "" Then
                    .Fields("Estado") = Me.Estado
                End If
                If Me.FechaEstado <> "" Then
                    .Fields("FechaEstado") = Me.FechaEstado
                End If
                If IsNumeric(Me.Priorizacion) Then
                    .Fields("Priorizacion") = Me.Priorizacion
                End If
                If IsDate(Me.FechaMaterializado) Then
                    .Fields("FechaMaterializado") = Me.FechaMaterializado
                End If
                If IsDate(Me.FechaRetirado) Then
                    .Fields("FechaRetirado") = Me.FechaRetirado
                End If
                If IsDate(Me.FechaCerrado) Then
                    .Fields("FechaCerrado") = Me.FechaCerrado
                End If
                If Me.Origen <> "" Then
                    .Fields("Origen") = Me.Origen
                End If
                If IsDate(Me.FechaMitigacionAceptar) Then
                    .Fields("FechaMitigacionAceptar") = Me.FechaMitigacionAceptar
                End If
                If Me.JustificacionAceptacionRiesgo <> "" Then
                    .Fields("JustificacionAceptacionRiesgo") = Me.JustificacionAceptacionRiesgo
                End If
                If IsDate(Me.FechaJustificacionAceptacionRiesgo) Then
                    .Fields("FechaJustificacionAceptacionRiesgo") = Me.FechaJustificacionAceptacionRiesgo
                End If
                If IsDate(Me.FechaAprobacionAceptacionPorCalidad) Then
                    .Fields("FechaAprobacionAceptacionPorCalidad") = Me.FechaAprobacionAceptacionPorCalidad
                End If
                If IsDate(Me.FechaRechazoAceptacionPorCalidad) Then
                    .Fields("FechaRechazoAceptacionPorCalidad") = Me.FechaRechazoAceptacionPorCalidad
                End If
                If Me.JustificacionRetiroRiesgo <> "" Then
                    .Fields("JustificacionRetiroRiesgo") = Me.JustificacionRetiroRiesgo
                End If
                If IsDate(Me.FechaJustificacionRetiroRiesgo) Then
                    .Fields("FechaJustificacionRetiroRiesgo") = Me.FechaJustificacionRetiroRiesgo
                End If
                If IsDate(Me.FechaAprobacionRetiroPorCalidad) Then
                    .Fields("FechaAprobacionRetiroPorCalidad") = Me.FechaAprobacionRetiroPorCalidad
                End If
                If IsDate(Me.FechaRechazoRetiroPorCalidad) Then
                    .Fields("FechaRechazoRetiroPorCalidad") = Me.FechaRechazoRetiroPorCalidad
                End If
                .Fields("RequiereRiesgoDeBiblioteca") = Me.RequiereRiesgoDeBiblioteca
                If Me.RequiereRiesgoDeBiblioteca = "Sí" Then
                    If Me.CodRiesgoBiblioteca <> "" Then
                        .Fields("CodRiesgoBiblioteca") = Me.CodRiesgoBiblioteca
                        If Me.CodRiesgoBiblioteca = m_ObjEntorno.CodigoBibliotecaParaRetipificar Then
                            Me.RiesgoPendienteRetipificacion = "Sí"
                           
                            .Fields("RiesgoPendienteRetipificacion") = Me.RiesgoPendienteRetipificacion
                            Me.FechaRiesgoParaRetipificar = CStr(m_FechaRef)
                            .Fields("FechaRiesgoParaRetipificar") = Me.FechaRiesgoParaRetipificar
                             m_ParaRetipificarAhora = EnumSiNo.Sí
                        Else
                            m_ParaRetipificarAhora = EnumSiNo.No
                        End If
                    Else
                        m_ParaRetipificarAntes = EnumSiNo.No
                    End If
                Else
                    m_ParaRetipificarAhora = EnumSiNo.No
                End If
                Me.HayErrorEnRiesgo = Me.HayErrorEnRiesgoCalculado
                .Fields("HayErrorEnRiesgo") = Me.HayErrorEnRiesgo
                Me.NombreNodoDesc = Me.NombreNodoDescCalculado
                .Fields("NombreNodoDesc") = Me.NombreNodoDesc
                Me.NombreNodoEstado = Me.NombreNodoEstadoCalculado
                .Fields("NombreNodoEstado") = Me.NombreNodoEstado
                
               
            .Update
        Else
            .Edit
                
                .Fields("CodigoUnico") = Me.CodigoUnico
                .Fields("CodigoRiesgo") = Me.CodigoRiesgo
                .Fields("FechaDetectado") = Me.FechaDetectado
                .Fields("DetectadoPor") = Me.DetectadoPor
                .Fields("EntidadDetecta") = Me.EntidadDetecta
                .Fields("Plazo") = Me.Plazo
                .Fields("Calidad") = Me.Calidad
                .Fields("Coste") = Me.Coste
                .Fields("ImpactoGlobal") = Me.ImpactoGlobal
                .Fields("Vulnerabilidad") = Me.Vulnerabilidad
                .Fields("Valoracion") = Me.Valoracion
                .Fields("Mitigacion") = Me.Mitigacion
                .Fields("Contingencia") = Me.Contingencia
                .Fields("RequierePlanContingencia") = Me.RequierePlanContingencia
                .Fields("Descripcion") = Me.Descripcion
                If Me.CausaRaiz <> "" Then
                    .Fields("CausaRaiz") = Me.CausaRaiz
                Else
                    .Fields("CausaRaiz") = Null
                End If
                .Fields("Estado") = Me.Estado
                .Fields("FechaEstado") = CStr(m_FechaRef)
                If IsNumeric(Me.Priorizacion) Then
                    .Fields("Priorizacion") = Me.Priorizacion
                Else
                    .Fields("Priorizacion") = Null
                End If
                If IsDate(Me.FechaMaterializado) Then
                    .Fields("FechaMaterializado") = Me.FechaMaterializado
                Else
                    .Fields("FechaMaterializado") = Null
                End If
                If IsDate(Me.FechaRetirado) Then
                    .Fields("FechaRetirado") = Me.FechaRetirado
                Else
                    .Fields("FechaRetirado") = Null
                End If
                If IsDate(Me.FechaCerrado) Then
                    .Fields("FechaCerrado") = Me.FechaCerrado
                Else
                    .Fields("FechaCerrado") = Null
                End If
                If IsDate(Me.FechaMitigacionAceptar) Then
                    .Fields("FechaMitigacionAceptar") = Me.FechaMitigacionAceptar
                Else
                    .Fields("FechaMitigacionAceptar") = Null
                End If
                If Me.JustificacionAceptacionRiesgo <> "" Then
                    .Fields("JustificacionAceptacionRiesgo") = Me.JustificacionAceptacionRiesgo
                Else
                    .Fields("JustificacionAceptacionRiesgo") = Null
                End If
                If IsDate(Me.FechaJustificacionAceptacionRiesgo) Then
                    .Fields("FechaJustificacionAceptacionRiesgo") = Me.FechaJustificacionAceptacionRiesgo
                Else
                    .Fields("FechaJustificacionAceptacionRiesgo") = Null
                End If
                If IsDate(Me.FechaAprobacionAceptacionPorCalidad) Then
                    .Fields("FechaAprobacionAceptacionPorCalidad") = Me.FechaAprobacionAceptacionPorCalidad
                Else
                    .Fields("FechaAprobacionAceptacionPorCalidad") = Null
                End If
                If IsDate(Me.FechaRechazoAceptacionPorCalidad) Then
                    .Fields("FechaRechazoAceptacionPorCalidad") = Me.FechaRechazoAceptacionPorCalidad
                Else
                    .Fields("FechaRechazoAceptacionPorCalidad") = Null
                End If
                If Me.JustificacionRetiroRiesgo <> "" Then
                    .Fields("JustificacionRetiroRiesgo") = Me.JustificacionRetiroRiesgo
                Else
                    .Fields("JustificacionRetiroRiesgo") = Null
                End If
                If IsDate(Me.FechaJustificacionRetiroRiesgo) Then
                    .Fields("FechaJustificacionRetiroRiesgo") = Me.FechaJustificacionRetiroRiesgo
                Else
                    .Fields("FechaJustificacionRetiroRiesgo") = Null
                End If
                If IsDate(Me.FechaAprobacionRetiroPorCalidad) Then
                    .Fields("FechaAprobacionRetiroPorCalidad") = Me.FechaAprobacionRetiroPorCalidad
                Else
                    .Fields("FechaAprobacionRetiroPorCalidad") = Null
                End If
                If IsDate(Me.FechaRechazoRetiroPorCalidad) Then
                    .Fields("FechaRechazoRetiroPorCalidad") = Me.FechaRechazoRetiroPorCalidad
                Else
                    .Fields("FechaRechazoRetiroPorCalidad") = Null
                End If
                
                 .Fields("RequiereRiesgoDeBiblioteca") = Me.RequiereRiesgoDeBiblioteca
                If Me.CodRiesgoBiblioteca <> "" Then
                    .Fields("CodRiesgoBiblioteca") = Me.CodRiesgoBiblioteca
                End If
                If Me.Origen <> "" Then
                    .Fields("Origen") = Me.Origen
                Else
                    .Fields("Origen") = Null
                End If
                
                If Me.RequiereRiesgoDeBiblioteca = "Sí" Then
                    If Me.CodRiesgoBiblioteca = m_ObjEntorno.CodigoBibliotecaParaRetipificar Then
                        .Fields("CodRiesgoBiblioteca") = Me.CodRiesgoBiblioteca
                        Me.RiesgoPendienteRetipificacion = "Sí"
                        .Fields("RiesgoPendienteRetipificacion") = Me.RiesgoPendienteRetipificacion
                        Me.FechaRiesgoParaRetipificar = CStr(m_FechaRef)
                        .Fields("FechaRiesgoParaRetipificar") = Me.FechaRiesgoParaRetipificar
                        
                    End If
                    If p_ObjRiesgoAlInicio.CodRiesgoBiblioteca = m_ObjEntorno.CodigoBibliotecaParaRetipificar Then
                        'aquí ya no es codigo de bi
                        Me.FechaRiesgoRetipificado = CStr(m_FechaRef)
                        Me.RiesgoPendienteRetipificacion = "No"
                        Me.FechaRiesgoParaRetipificar = ""
                        .Fields("CodRiesgoBiblioteca") = Me.CodRiesgoBiblioteca
                        .Fields("FechaRiesgoRetipificado") = Me.FechaRiesgoRetipificado
                        .Fields("FechaRiesgoParaRetipificar") = Null
                        .Fields("RiesgoPendienteRetipificacion") = "No"
                    End If
                    If Me.CodRiesgoBiblioteca = m_ObjEntorno.CodigoBibliotecaParaRetipificar Then
                        m_ParaRetipificarAhora = EnumSiNo.Sí
                    Else
                        m_ParaRetipificarAhora = EnumSiNo.No
                    End If
                    If p_ObjRiesgoAlInicio.CodRiesgoBiblioteca = m_ObjEntorno.CodigoBibliotecaParaRetipificar Then
                        m_ParaRetipificarAntes = EnumSiNo.Sí
                    Else
                        m_ParaRetipificarAntes = EnumSiNo.No
                    End If
                Else
                    Me.CodRiesgoBiblioteca = ""
                    Me.FechaRiesgoParaRetipificar = ""
                    Me.FechaRiesgoRetipificado = ""
                    Me.DiasSinRespuestaCalidadRetipificacion = ""
                    .Fields("CodRiesgoBiblioteca") = Null
                    .Fields("FechaRiesgoParaRetipificar") = Null
                    .Fields("FechaRiesgoRetipificado") = Null
                    .Fields("DiasSinRespuestaCalidadRetipificacion") = Null
                    
                End If
                Me.HayErrorEnRiesgo = Me.HayErrorEnRiesgoCalculado
                .Fields("HayErrorEnRiesgo") = Me.HayErrorEnRiesgo
                Me.NombreNodoDesc = Me.NombreNodoDescCalculado
                .Fields("NombreNodoDesc") = Me.NombreNodoDesc
                Me.NombreNodoEstado = Me.NombreNodoEstadoCalculado
                .Fields("NombreNodoEstado") = Me.NombreNodoEstado
                 
                
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If p_ObjRiesgoAlInicio Is Nothing Then
        If m_ParaRetipificarAhora = EnumSiNo.Sí Then
            EnviarCorreoRiesgoRequiereRetipificacion Me, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    Else
        If m_ParaRetipificarAntes = EnumSiNo.No Then
            If m_ParaRetipificarAhora = EnumSiNo.Sí Then
                EnviarCorreoRiesgoRequiereRetipificacion Me, p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If EsTecnico <> EnumSiNo.Sí Then
                    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                End If
            End If
        Else
            If m_ParaRetipificarAhora = EnumSiNo.No Then
                If EsTecnico <> EnumSiNo.Sí Then
                    'correo al técnico de que ya está tipificado
                    EnviarCorreoRiesgoRetipificacion Me, p_ObjRiesgoAlInicio, p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                End If
            End If
        End If
    End If
    RegistrarActualizacionRiesgo p_Riesgo:=Me, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionEdicion p_Edicion:=Me.Edicion, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_Comprobando = EnumSiNo.No Then
        Exit Function
    End If
    Dim m_EstadoAlInicio As EnumRiesgoEstado
    With Me
        If Me.Edicion.Proyecto.ParaInformeAvisos <> "No" And Not p_ObjRiesgoAlInicio Is Nothing Then
            
            m_EstadoAlInicio = p_ObjRiesgoAlInicio.EstadoEnum
            If .Mitigacion = "Aceptar" Then
                
                If p_ObjRiesgoAlInicio Is Nothing Then
                    EnviarCorreoTecnicoRiesgoAceptado Me, p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Else
                    
                    If m_EstadoAlInicio <> EnumRiesgoEstado.AceptadoSinVisar Then
                        EnviarCorreoTecnicoRiesgoAceptado Me, p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    End If
                End If
            End If
            If IsDate(Nz(Me.FechaRetirado, "")) Then
                
                If p_ObjRiesgoAlInicio Is Nothing Then
                    EnviarCorreoTecnicoRiesgoRetirado Me, p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Else
                    
                    If m_EstadoAlInicio <> EnumRiesgoEstado.RetiradoSinVisar Then
                        EnviarCorreoTecnicoRiesgoRetirado Me, p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    End If
                End If
            
            End If
            If IsDate(.FechaMaterializado) Then
                If p_ObjRiesgoAlInicio Is Nothing Then
                    EnviarCorreoRiesgoMaterializado Me, p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    .RegistrarRiesgoMaterializado Me.Estado, p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Else
                    If Not IsDate(p_ObjRiesgoAlInicio.FechaMaterializado) Then
                        EnviarCorreoRiesgoMaterializado Me, p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    End If
                End If
            Else
                If Not p_ObjRiesgoAlInicio Is Nothing Then
                    If IsDate(p_ObjRiesgoAlInicio.FechaMaterializado) Then
                        .RegistrarRiesgoMaterializado Me.Estado, p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    End If
                End If
            End If
        End If
        
    End With
    
    
    With Me.Edicion
        Set .colRiesgos = Nothing
        Set .ColRiesgosOferta = Nothing
        Set .ColRiesgosPorPrioridadTodos = Nothing
        
    End With
    With Me.Edicion.Proyecto
        Set .colEdiciones(Me.IDEdicion) = Me
        Set .ColRiesgosDeJP = Nothing
        Set .ColRiesgosExternos = Nothing
        Set .ColRiesgosOferta = Nothing
    End With
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function AltaDesdeProyecto( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                        
    
    Dim m_NombreResponsableCalidad As String
    Dim m_ObjProyecto As Proyecto
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    Me.CodigoRiesgo = Me.CodigoRiesgoCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.CodigoUnico = Me.CodigoUnicoCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjProyectoActivo.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
        Me.Priorizacion = Me.PriorizacionCalculado
    End If
    
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjProyecto = Me.Edicion.Proyecto
    m_NombreResponsableCalidad = m_ObjProyecto.NombreUsuarioCalidad
    If m_NombreResponsableCalidad = "" Then
        p_Error = "No se conoce el Usuario de Calidad que le va a hacer el seguimiento"
        Err.Raise 1000
    End If
    With Me
        .CodigoRiesgo = .CodigoRiesgoCalculado
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        .CodigoUnico = .CodigoUnicoCalculado
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        .EntidadDetecta = "Persona"
        If .Estado = "" Then
            .Estado = .ESTADOCalculadoTexto
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        .Estado = "Detectado"
        .FechaEstado = CStr(m_FechaRef)
        If m_ObjProyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
            .Priorizacion = Right(.CodigoRiesgo, 1)
        End If
        If m_NombreResponsableCalidad <> "" Then
            .DetectadoPor = m_NombreResponsableCalidad
        End If
        Me.IDRiesgo = Me.IDRiesgoCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    Registrar EnumSiNo.No, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AltaDesdeProyecto ha devuelto el error: " & Err.Description
    End If
End Function

Public Function AltaDesdeRiesgoExterno( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
    On Error GoTo errores
    
    Me.Error = ""
    If Me.CodigoRiesgo = "" Then
        Me.CodigoRiesgo = Me.CodigoRiesgoCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Registrar EnumSiNo.No, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.AltaDesdeRiesgoExterno ha devuelto el error: " & Err.Description
    End If
End Function




Public Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String

    Dim m_ID As Variant
    Dim m_ObjAnexo As Anexo
    Dim m_Borrable As EnumSiNo
    
    On Error GoTo errores
    
    m_Borrable = Me.Borrable
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Borrable = EnumSiNo.No Then
        p_Error = "Para que un riesgo se pueda borrar ha de cumplir estas condiciones:" & vbNewLine & _
                vbTab & "- edición activa(Sin fecha de publicación)" & vbNewLine & _
                vbTab & "-El Riesgo ha de haber nacido en esta edición(Sin fecha de publicación)" & vbNewLine & _
                vbTab & "-Riesgo no proviniente de un Alta de Calidad" & vbNewLine & _
                vbTab & "-Ha de tener permisos sobre la gestión de riegos a la que pertenece el mismo"
        Err.Raise 1000
    End If
    Set Me.ColAnexos = Nothing
    If Not Me.ColAnexos Is Nothing Then
        For Each m_ID In Me.ColAnexos
            Set m_ObjAnexo = Me.ColAnexos(m_ID)
            With m_ObjAnexo
                If fso.FileExists(.URLAnexo) Then
                    If FicheroAbierto(.URLAnexo) Then
                        p_Error = "Al menos un anexo del riesgo está abierto"
                        Err.Raise 1000
                    End If
                End If
            End With
            Set m_ObjAnexo = Nothing
        Next
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function



Public Function AceptacionAprobar( _
                                    p_Fecha As String, _
                                    Optional ByRef p_Estado As EnumRiesgoEstado, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    If Me.IDRiesgo = "" Then
        p_Error = "El riesgo no está registrado"
        Err.Raise 1000
    End If
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "El Riesgo no pertence a una edición activa"
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        p_Error = "Operación solo autorizada a Calidad"
        Err.Raise 1000
    End If
    If p_Estado = Empty Then
        p_Estado = Me.ESTADOCalculado
    End If
    If p_Estado <> EnumRiesgoEstado.AceptadoSinVisar Then
        p_Error = "La Aceptación del Riesgo no aparece en estado aceptado sin visar"
        Err.Raise 1000
    End If
    Me.FechaAprobacionAceptacionPorCalidad = p_Fecha
    Me.FechaRechazoAceptacionPorCalidad = ""
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaAprobacionAceptacionPorCalidad") = Me.FechaAprobacionAceptacionPorCalidad
            .Fields("FechaRechazoAceptacionPorCalidad") = Null
            .Fields("DiasSinRespuestaCalidadRetiro") = Null
            Me.Estado = "Aceptado"
            .Fields("Estado") = Me.Estado
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    p_Estado = EnumRiesgoEstado.Aceptado
    If Me.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
        EnviarCorreoCalidadApruebaRiesgoAceptado Me, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.AceptacionAprobar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function AceptacionAprobarQuitar( _
                                        Optional ByRef p_Estado As EnumRiesgoEstado, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    
    If Me.IDRiesgo = "" Then
        p_Error = "El riesgo no está registrado"
        Err.Raise 1000
    End If
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "El Riesgo no pertence a una edición activa"
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        p_Error = "Operación solo autorizada a Calidad"
        Err.Raise 1000
    End If
    If p_Estado = Empty Then
        p_Estado = Me.ESTADOCalculado
    End If
    If p_Estado <> EnumRiesgoEstado.Aceptado Then
        p_Error = "El Riesgo no aparece en estado aceptado"
        Err.Raise 1000
    End If
    Me.FechaAprobacionAceptacionPorCalidad = ""
    Me.FechaRechazoAceptacionPorCalidad = ""
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaAprobacionAceptacionPorCalidad") = Null
            .Fields("FechaRechazoAceptacionPorCalidad") = Null
            .Fields("DiasSinRespuestaCalidadRetiro") = Null
            Me.Estado = "AceptadoSinVisar"
            .Fields("Estado") = Me.Estado
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    p_Estado = EnumRiesgoEstado.AceptadoSinVisar
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.AceptacionAprobarQuitar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function AceptacionRegistrar( _
                                    p_Justificacion As String, _
                                    Optional p_Fecha As String, _
                                    Optional ByRef p_Estado As EnumRiesgoEstado, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_RiesgoEnRetirada As EnumSiNo
    On Error GoTo errores
    
    Me.Error = ""
    If p_Justificacion = "" Then
        p_Error = "Se ha de justificar la aceptación"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    With Me
        If .EsActivo = EnumSiNo.No Then
            p_Error = "El riesgo no es de una edición activa"
            Err.Raise 1000
        End If
        If .UsuarioConectadoAutorizado = EnumSiNo.No Then
            p_Error = "Usuario no autorizado"
            Err.Raise 1000
        End If
        If p_Estado = Empty Then
            p_Estado = .ESTADOCalculado
        End If
        
        m_RiesgoEnRetirada = RiesgoEnRetirada(p_Estado, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_RiesgoEnRetirada = EnumSiNo.Sí Then
            p_Error = "El Riesgo está Retirado o en proceso"
            Err.Raise 1000
        End If
        If p_Estado = EnumRiesgoEstado.Materializado Then
            p_Error = "El Riesgo ya aparece como Materializado"
            Err.Raise 1000
        End If
        If p_Estado = EnumRiesgoEstado.Aceptado Then
            p_Error = "El Riesgo ya aparece como Aceptado y visado por calidad"
            Err.Raise 1000
        End If
        If p_Estado = EnumRiesgoEstado.AceptadoRechazado Then
            If p_Justificacion = .JustificacionAceptacionRiesgo Then
                p_Error = "El Riesgo está rechazado y no ha cambiado la justificación"
                Err.Raise 1000
            End If
        End If
        If .JustificacionAceptacionRiesgo <> "" Then
            If p_Justificacion = .JustificacionAceptacionRiesgo Then
                p_Error = "El Riesgo ya estaba pendiente de justificar y no se ha cambiado la misma"
                Err.Raise 1000
            End If
        End If
        .FechaRechazoAceptacionPorCalidad = ""
        .JustificacionAceptacionRiesgo = p_Justificacion
        .FechaJustificacionAceptacionRiesgo = p_Fecha
        .FechaMitigacionAceptar = p_Fecha
        .Mitigacion = "Aceptar"
    End With
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaRechazoAceptacionPorCalidad") = Null
            .Fields("JustificacionAceptacionRiesgo") = Me.JustificacionAceptacionRiesgo
            .Fields("FechaJustificacionAceptacionRiesgo") = Me.FechaJustificacionAceptacionRiesgo
            .Fields("FechaMitigacionAceptar") = Me.FechaMitigacionAceptar
            .Fields("Mitigacion") = Me.Mitigacion
            Me.Estado = "AceptadoSinVisar"
            .Fields("Estado") = Me.Estado
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    p_Estado = EnumRiesgoEstado.AceptadoSinVisar
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
        EnviarCorreoTecnicoRiesgoAceptado Me, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.AceptacionRegistrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function AceptacionRegistrarQuitar( _
                                            p_Mitigacion As String, _
                                            Optional p_Fecha As String, _
                                            Optional ByRef p_Estado As EnumRiesgoEstado, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_RiesgoEnAceptacion As EnumSiNo
    Dim m_RiesgoEnRetirada As EnumSiNo
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    With Me
        If .EsActivo = EnumSiNo.No Then
            p_Error = "El riesgo no es de una edición activa"
            Err.Raise 1000
        End If
        If .UsuarioConectadoAutorizado = EnumSiNo.No Then
            p_Error = "Usuario no autorizado"
            Err.Raise 1000
        End If
        If p_Estado = Empty Then
            p_Estado = Me.ESTADOCalculado
        End If
        
        m_RiesgoEnAceptacion = RiesgoEnAceptacion(p_Estado)
        If m_RiesgoEnAceptacion = EnumSiNo.No Then
            p_Error = "El Riesgo no está en proceso de Aceptación"
            Err.Raise 1000
        End If
        m_RiesgoEnRetirada = RiesgoEnRetirada(p_Estado)
        If m_RiesgoEnRetirada = EnumSiNo.Sí Then
            p_Error = "El Riesgo está en proceso de retirada"
            Err.Raise 1000
        End If
        If p_Estado = EnumRiesgoEstado.Materializado Then
            p_Error = "El Riesgo está materializado"
            Err.Raise 1000
        End If
        If p_Mitigacion = "Aceptar" Then
            p_Error = "La nueva Mitigación no puede ser Aceptar"
            Err.Raise 1000
        End If
        .FechaRechazoAceptacionPorCalidad = ""
        .JustificacionAceptacionRiesgo = ""
        .FechaJustificacionAceptacionRiesgo = ""
        .FechaMitigacionAceptar = ""
        .Mitigacion = p_Mitigacion
    End With
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaRechazoAceptacionPorCalidad") = Null
            .Fields("JustificacionAceptacionRiesgo") = Null
            .Fields("FechaJustificacionAceptacionRiesgo") = Null
            .Fields("FechaMitigacionAceptar") = Null
            .Fields("Mitigacion") = Me.Mitigacion
            Me.Estado = Me.ESTADOCalculadoTexto
            .Fields("Estado") = Me.Estado
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    p_Estado = Me.EstadoEnum
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.AceptacionRegistrarQuitar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function RetiroAprobar( _
                                Optional p_Fecha As String, _
                                Optional ByRef p_Estado As EnumRiesgoEstado, _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    If Me.IDRiesgo = "" Then
        p_Error = "El riesgo no está registrado"
        Err.Raise 1000
    End If
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "El Riesgo no pertence a una edición activa"
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        p_Error = "Operación solo autorizada a Calidad"
        Err.Raise 1000
    End If
    If p_Estado = Empty Then
        p_Estado = Me.ESTADOCalculado
    End If
    If p_Estado <> EnumRiesgoEstado.RetiradoSinVisar Then
        p_Error = "El retiro del Riesgo no aparece en estado retirado sin visar"
        Err.Raise 1000
    End If
    Me.FechaAprobacionRetiroPorCalidad = p_Fecha
    Me.FechaRechazoRetiroPorCalidad = ""
    
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaAprobacionRetiroPorCalidad") = Me.FechaAprobacionRetiroPorCalidad
            .Fields("FechaRechazoRetiroPorCalidad") = Null
            Me.Estado = "Retirado"
            .Fields("Estado") = Me.Estado
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Me.Cerrar p_Fecha, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    p_Estado = EnumRiesgoEstado.Retirado
    If Me.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
        EnviarCorreoCalidadApruebaRiesgoRetirado Me, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.RetiroAprobar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function RetiroAprobacionQuitar( _
                                        Optional ByRef p_Estado As EnumRiesgoEstado, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    
    If Me.IDRiesgo = "" Then
        p_Error = "El riesgo no está registrado"
        Err.Raise 1000
    End If
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "El Riesgo no pertence a una edición activa"
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        p_Error = "La operación sólo la puede hacer un miembro de Calidad"
        Err.Raise 1000
    End If
    If p_Estado = Empty Then
        p_Estado = Me.ESTADOCalculado
    End If
    If p_Estado <> EnumRiesgoEstado.Retirado Then
        p_Error = "El Riesgo no aparece en estado retirado"
        Err.Raise 1000
    End If
    Me.FechaRechazoRetiroPorCalidad = ""
    Me.FechaAprobacionRetiroPorCalidad = ""
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaRechazoRetiroPorCalidad") = Null
            .Fields("FechaAprobacionRetiroPorCalidad") = Null
            Me.Estado = "RetiradoSinVisar"
            .Fields("Estado") = Me.Estado
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    p_Estado = EnumRiesgoEstado.RetiradoSinVisar
    If Me.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
        EnviarCorreoCalidadQuitarAprobacionRiesgoRetirado Me, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.RetiroAprobacionQuitar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function RetiroRegistrar( _
                                    p_Justificacion As String, _
                                    Optional p_Fecha As String, _
                                    Optional ByRef p_Estado As EnumRiesgoEstado, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim rcdDatos As DAO.Recordset
    
    Dim m_RiesgoEnAceptacion As EnumSiNo
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    If p_Justificacion = "" Then
        p_Error = "Se ha de justificar la retirada del riesgo"
        Err.Raise 1000
    End If
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If
    With Me
        If .EsActivo = EnumSiNo.No Then
            p_Error = "El riesgo no es de una edición activa"
            Err.Raise 1000
        End If
        If .UsuarioConectadoAutorizado = EnumSiNo.No Then
            p_Error = "EL usuario no está autorizado"
            Err.Raise 1000
        End If
        If p_Estado = Empty Then
            p_Estado = Me.ESTADOCalculado
        End If
        If p_Estado = EnumRiesgoEstado.RetiradoSinVisar Then
            If p_Justificacion = .JustificacionRetiroRiesgo Then
                p_Error = "El riesgo ya estaba en proceso de retirada y no se ha cambiado la justificación"
                Err.Raise 1000
            End If
        End If
        m_RiesgoEnAceptacion = RiesgoEnAceptacion(p_Estado)
        If m_RiesgoEnAceptacion = EnumSiNo.Sí Then
            p_Error = "El Riesgo está en proceso de Aceptación"
            Err.Raise 1000
        End If
        If p_Estado = EnumRiesgoEstado.Materializado Then
            p_Error = "El Riesgo ya aparece como Materializado"
            Err.Raise 1000
        End If
        .FechaRetirado = p_Fecha
        .JustificacionRetiroRiesgo = p_Justificacion
        .FechaJustificacionRetiroRiesgo = p_Fecha
    End With
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaRetirado") = p_Fecha
            .Fields("JustificacionRetiroRiesgo") = p_Justificacion
            .Fields("FechaJustificacionRetiroRiesgo") = p_Fecha
            Me.Estado = "RetiradoSinVisar"
            .Fields("Estado") = Me.Estado
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    p_Estado = EnumRiesgoEstado.RetiradoSinVisar
    If Me.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
        EnviarCorreoTecnicoRiesgoRetirado Me, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    RegistrarActualizacionRiesgo p_Riesgo:=Me, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionEdicion p_Edicion:=Me.Edicion, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.RetiroRegistrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function RetiroRegistrarQuitar( _
                                        Optional ByRef p_Estado As EnumRiesgoEstado, _
                                        Optional ByRef p_Error As String _
                                    ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    
    With Me
        If Me.IDRiesgo = "" Then
            p_Error = "El riesgo no está registrado"
            Err.Raise 1000
        End If
        If Me.EsActivo = EnumSiNo.No Then
            p_Error = "El Riesgo no pertence a una edición activa"
            Err.Raise 1000
        End If
        If .UsuarioConectadoAutorizado = EnumSiNo.No Then
            p_Error = "EL usuario no está autorizado"
            Err.Raise 1000
        End If
        If p_Estado = Empty Then
            p_Estado = Me.ESTADOCalculado
        End If
'        If p_Estado <> EnumRiesgoEstado.RetiradoSinVisar And p_Estado <> EnumRiesgoEstado.RetiradoRechazado Then
'            p_Error = "El Retiro del Riesgo ha de estar sin visar por Calidad o Rechazado por ellos"
'            Err.Raise 1000
'        End If
        .FechaRetirado = ""
        .JustificacionRetiroRiesgo = ""
        .FechaJustificacionRetiroRiesgo = ""
        .FechaAprobacionRetiroPorCalidad = ""
        .FechaRechazoRetiroPorCalidad = ""
    End With
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaRetirado") = Null
            .Fields("JustificacionRetiroRiesgo") = Null
            .Fields("FechaJustificacionRetiroRiesgo") = Null
            .Fields("FechaAprobacionRetiroPorCalidad") = Null
            .Fields("FechaRechazoRetiroPorCalidad") = Null
            .Fields("DiasSinRespuestaCalidadRetiro") = Null
            
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Me.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
        EnviarCorreoCalidadQuitarRechazoRiesgoRetirado Me, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    EstadoRegistrar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionRiesgo p_Riesgo:=Me, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionEdicion p_Edicion:=Me.Edicion, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.RetiroRegistrarQuitar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function ValidarFechaMaterializacionPermitida( _
                                                    ByVal p_Fecha As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Boolean
    On Error GoTo errores

    p_Error = ""

    If Not IsDate(p_Fecha) Then
        ValidarFechaMaterializacionPermitida = True
        Exit Function
    End If

    If DateValue(CDate(p_Fecha)) > Date Then
        p_Error = "La fecha de materialización no puede ser posterior a la fecha actual"
        ValidarFechaMaterializacionPermitida = False
        Exit Function
    End If

    ValidarFechaMaterializacionPermitida = True
    Exit Function
errores:
    p_Error = "El método Riesgo.ValidarFechaMaterializacionPermitida ha devuelto el error: " & Err.Description
    ValidarFechaMaterializacionPermitida = False
End Function
Public Function MaterializacionRegistrar( _
                                            Optional p_Fecha As String, _
                                            Optional ByRef p_Error As String, _
                                            Optional ByVal p_IDPlanContingencia As String _
                                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_RiesgoEnAceptacion As EnumSiNo
    Dim m_RiesgoEnRetirada As EnumSiNo
    On Error GoTo errores
    
    Me.Error = ""
   
    If Not IsDate(p_Fecha) Then
        p_Fecha = Date
    End If

    If Not ValidarFechaMaterializacionPermitida(p_Fecha, p_Error) Then
        Err.Raise 1000
    End If

    If p_IDPlanContingencia <> "" Then
        If ValidarYResolverPlanContingenciaMaterializacion( _
            Me.IDRiesgo, _
            p_IDPlanContingencia, _
            p_Error) <> EnumSiNo.Sí Then
            Err.Raise 1000
        End If
    End If

    With Me
        If .EsActivo = EnumSiNo.No Then
            p_Error = "El riesgo no es de una edición activa"
            Err.Raise 1000
        End If
        
        m_RiesgoEnAceptacion = RiesgoEnAceptacion(Me.EstadoEnum)
        If m_RiesgoEnAceptacion = EnumSiNo.Sí Then
            p_Error = "El Riesgo está en proceso de Aceptación"
            Err.Raise 1000
        End If
        m_RiesgoEnRetirada = RiesgoEnRetirada(Me.EstadoEnum)
        If m_RiesgoEnRetirada = EnumSiNo.Sí Then
            p_Error = "El Riesgo está en proceso de retirada"
            Err.Raise 1000
        End If
        .FechaMaterializado = p_Fecha
    End With
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaMaterializado") = Me.FechaMaterializado
            Me.Estado = "Materializado"
            .Fields("Estado") = Me.Estado
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
   
    If Me.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
        EnviarCorreoRiesgoMaterializado Me, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    RegistrarRiesgoMaterializado Me.Estado, p_Error, p_IDPlanContingencia
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionRiesgo p_Riesgo:=Me, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionEdicion p_Edicion:=Me.Edicion, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.MaterializacionRegistrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function MaterializacionQuitarRegistrar( _
                                                Optional ByRef p_Estado As EnumRiesgoEstado, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
   
    
    With Me
        If .EsActivo = EnumSiNo.No Then
            p_Error = "El riesgo no es de una edición activa"
            Err.Raise 1000
        End If
        If p_Estado = Empty Then
            p_Estado = .ESTADOCalculado
        End If
       
        If p_Estado <> EnumRiesgoEstado.Materializado Then
            p_Error = "El Riesgo no aparece como materializado"
            Err.Raise 1000
        End If
        
        .FechaMaterializado = ""
       
    End With
    m_SQL = "SELECT * " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & Me.IDRiesgo & ";"
    InvalidarCacheRiesgo Me.IDRiesgo, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("FechaMaterializado") = Null
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    EstadoRegistrar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarRiesgoMaterializado Me.Estado, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionRiesgo p_Riesgo:=Me, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarActualizacionEdicion p_Edicion:=Me.Edicion, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.MaterializacionQuitarRegistrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EstadoRegistrar( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_Estado As EnumRiesgoEstado
    Dim m_EstadoTexto As String
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Estado = Me.ESTADOCalculado
    m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
    If m_EstadoTexto = "" Then
        p_Error = "Estado no reconocido"
        Err.Raise 1000
    End If
    With Me
        
          m_SQL = "UPDATE TbRiesgos " & _
                "SET Estado='" & m_EstadoTexto & "' " & _
                "WHERE IDRiesgo=" & .IDRiesgo & ";"

        InvalidarCacheRiesgo Me.IDRiesgo, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        getdb().Execute m_SQL
        .Estado = m_EstadoTexto
    
    End With
    If Not m_ObjRiesgoActivo Is Nothing Then
        If CStr(m_ObjRiesgoActivo.IDRiesgo) = CStr(Me.IDRiesgo) Then
            m_ObjRiesgoActivo.Estado = m_EstadoTexto
        End If
    End If




    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.EstadoRegistrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Cerrar( _
                        Optional p_FechaCierre As String, _
                        Optional ByRef p_Error As String _
                        ) As String
    '---------------------------------------------------------------------------------------------------------
    'Va a recorrer todos los planes de mitigación y de contingencia activos y los va a ir cerrando
    '---------------------------------------------------------------------------------------------------------
    Dim m_Col As Scripting.Dictionary
    
    Dim m_PCAccion As PCAccion
    Dim m_PCAccionAlInicio As PCAccion
    Dim m_PMAccion As PMAccion
    Dim m_PMAccionAlInicio As PMAccion
    Dim m_ID As Variant
    On Error GoTo errores
    If Not IsDate(p_FechaCierre) Then p_FechaCierre = Date
    Set m_Col = Me.colAccionesMitigacionSinCerrar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_PMAccion = m_Col(m_ID)
            Set m_PMAccionAlInicio = Constructor.getPMAccion(p_IDPMAccion:=CStr(m_ID), p_Error:=p_Error)
            
            If Not IsDate(m_PMAccion.FechaInicio) Then
                m_PMAccion.FechaInicio = p_FechaCierre
            End If
            If Not IsDate(m_PMAccion.FechaFinPrevista) Then
                m_PMAccion.FechaFinPrevista = p_FechaCierre
            End If
            If Not IsDate(m_PMAccion.FechaFinReal) Then
                m_PMAccion.FechaFinReal = p_FechaCierre
            End If
            m_PMAccion.Registrar m_PMAccionAlInicio, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_PMAccion = Nothing
        Next
        
    End If
    Set m_Col = Me.colAccionesContingenciaSinCerrar
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_PCAccion = m_Col(m_ID)
            Set m_PCAccionAlInicio = Constructor.getPCAccion(p_IDPCAccion:=CStr(m_ID), p_Error:=p_Error)
            If Not IsDate(m_PCAccion.FechaInicio) Then
                m_PCAccion.FechaInicio = p_FechaCierre
            End If
            If Not IsDate(m_PCAccion.FechaFinPrevista) Then
                m_PCAccion.FechaFinPrevista = p_FechaCierre
            End If
            If Not IsDate(m_PCAccion.FechaFinReal) Then
                m_PCAccion.FechaFinReal = p_FechaCierre
            End If
            m_PCAccion.Registrar m_PCAccionAlInicio, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_PCAccion = Nothing
        Next
        
    End If
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Cerrar ha devuelto el error: " & Err.Description
    End If
                                                
End Function



Public Function GrabarCamposCalculados(Optional ByRef p_Error As String) As String
    
    Dim m_ContingenciaCalculada As String
    Dim m_EstadoCalculadoTexto As String
    Dim m_RequiereContingenciaCalculadoTexto As String
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    With Me
        m_ContingenciaCalculada = .ContingenciaCalculada
        m_EstadoCalculadoTexto = .ESTADOCalculadoTexto
        m_RequiereContingenciaCalculadoTexto = .RequierePlanContingenciaCalculadoTexto
        
        m_SQL = "UPDATE TbRiesgos SET " & _
                "TbRiesgos.Contingencia = '" & m_ContingenciaCalculada & _
                "', TbRiesgos.RequierePlanContingencia = '" & m_RequiereContingenciaCalculadoTexto & _
                "', TbRiesgos.Estado = '" & m_EstadoCalculadoTexto & "' " & _
                "WHERE (((TbRiesgos.IDRiesgo)=" & Me.IDRiesgo & "));"
        InvalidarCacheRiesgo Me.IDRiesgo, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        getdb().Execute m_SQL
        .Contingencia = m_ContingenciaCalculada
        .Estado = m_EstadoCalculadoTexto
        .RequierePlanContingencia = m_RequiereContingenciaCalculadoTexto
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método GrabarCamposCalculados ha devuelto el error: " & vbNewLine & Err.Description
        
    End If
End Function

Public Function GrabarContingenciaCalculada(Optional ByRef p_Error As String) As String
    
    Dim m_ContingenciaCalculada As String
    Dim m_SQL As String
    
    
    On Error GoTo errores
    With Me
        m_ContingenciaCalculada = .ContingenciaCalculada
                
        m_SQL = "UPDATE TbRiesgos SET " & _
                "TbRiesgos.Contingencia = '" & m_ContingenciaCalculada & "' " & _
                "WHERE (((TbRiesgos.IDRiesgo)=" & Me.IDRiesgo & "));"
        InvalidarCacheRiesgo Me.IDRiesgo, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        getdb().Execute m_SQL
        .Contingencia = m_ContingenciaCalculada

    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método GrabarContingenciaCalculada ha devuelto el error: " & vbNewLine & Err.Description
        
    End If
End Function
Public Function GrabarEstadoCalculado(Optional ByRef p_Error As String) As String
    
    
    Dim m_EstadoCalculadoTexto As String
    Dim m_SQL As String
    
    
    On Error GoTo errores
    With Me
        m_EstadoCalculadoTexto = .ESTADOCalculadoTexto
        
        m_SQL = "UPDATE TbRiesgos SET " & _
                "TbRiesgos.Estado = '" & m_EstadoCalculadoTexto & "' " & _
                "WHERE (((TbRiesgos.IDRiesgo)=" & Me.IDRiesgo & "));"
        InvalidarCacheRiesgo Me.IDRiesgo, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        getdb().Execute m_SQL
        .Estado = m_EstadoCalculadoTexto

    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método GrabarEstadoCalculado ha devuelto el error: " & vbNewLine & Err.Description
        
    End If
End Function

Public Function GrabarRequiereContingenciaCalculado(Optional ByRef p_Error As String) As String
    
    
    Dim m_RequiereContingenciaCalculadoTexto As String
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    With Me
       
        m_RequiereContingenciaCalculadoTexto = .RequierePlanContingenciaCalculadoTexto
        
        m_SQL = "UPDATE TbRiesgos SET " & _
                "TbRiesgos.RequierePlanContingencia = '" & m_RequiereContingenciaCalculadoTexto & "' " & _
                "WHERE (((TbRiesgos.IDRiesgo)=" & Me.IDRiesgo & "));"
        InvalidarCacheRiesgo Me.IDRiesgo, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        getdb().Execute m_SQL

        .RequierePlanContingencia = m_RequiereContingenciaCalculadoTexto
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método GrabarRequiereContingenciaCalculado ha devuelto el error: " & vbNewLine & Err.Description
        
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
        .Add "IDRiesgo"
        .Add "IDEdicion"
        .Add "CodigoUnico"
        .Add "CodigoRiesgo"
        .Add "FechaDetectado"
        .Add "DetectadoPor"
        .Add "EntidadDetecta"
        .Add "Plazo"
        .Add "Calidad"
        .Add "Coste"
        .Add "ImpactoGlobal"
        .Add "Vulnerabilidad"
        .Add "Valoracion"
        .Add "Mitigacion"
        .Add "Contingencia"
        .Add "RequierePlanContingencia"
        .Add "Descripcion"
        .Add "CausaRaiz"
        .Add "Estado"
        .Add "FechaEstado"
        .Add "Priorizacion"
        .Add "FechaMaterializado"
        .Add "FechaRetirado"
        .Add "FechaCerrado"
        .Add "FechaMitigacionAceptar"
        .Add "JustificacionAceptacionRiesgo"
        .Add "FechaJustificacionAceptacionRiesgo"
        .Add "FechaAprobacionAceptacionPorCalidad"
        .Add "FechaRechazoAceptacionPorCalidad"
        .Add "JustificacionRetiroRiesgo"
        .Add "FechaJustificacionRetiroRiesgo"
        .Add "FechaAprobacionRetiroPorCalidad"
        .Add "FechaRechazoRetiroPorCalidad"
        .Add "RequiereRiesgoDeBiblioteca"
        .Add "CodRiesgoBiblioteca"
        .Add "RiesgoPendienteRetipificacion"
        .Add "FechaRiesgoParaRetipificar"
        .Add "FechaRiesgoRetipificado"
        .Add "DiasSinRespuestaCalidadAceptacion"
        .Add "DiasSinRespuestaCalidadRetiro"
        .Add "DiasSinRespuestaCalidadRetipificacion"
        .Add "Origen"
        .Add "HayErrorEnRiesgo"
        .Add "NombreNodoDesc"
        .Add "NombreNodoEstado"
        .Add "NombreIcono"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
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
        
        .Add "FechaDetectado"
        .Add "DetectadoPor"
        .Add "EntidadDetecta"
        .Add "Plazo"
        .Add "Calidad"
        .Add "Coste"
        .Add "ImpactoGlobal"
        .Add "Vulnerabilidad"
        .Add "Valoracion"
        .Add "Mitigacion"
        .Add "Contingencia"
        .Add "RequierePlanContingencia"
        .Add "Descripcion"
        .Add "CausaRaiz"
        .Add "Estado"
        .Add "Priorizacion"
        .Add "FechaMaterializado"
        .Add "FechaRetirado"
        .Add "FechaCerrado"
        .Add "CodRiesgoBiblioteca"
        
        
    End With
    Set ColCamposParaCambios = m_objColCamposParaCambios
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColCamposParaCambios ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDRiesgo" Then
        getPropiedad = Me.IDRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        getPropiedad = Me.IDEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoUnico" Then
        getPropiedad = Me.CodigoUnico
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoRiesgo" Then
        getPropiedad = Me.CodigoRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDetectado" Then
        getPropiedad = Me.FechaDetectado
        Exit Function
    End If
    If m_NombrePropiedad = "DetectadoPor" Then
        getPropiedad = Me.DetectadoPor
        Exit Function
    End If
    If m_NombrePropiedad = "EntidadDetecta" Then
        getPropiedad = Me.EntidadDetecta
        Exit Function
    End If
    If m_NombrePropiedad = "Plazo" Then
        getPropiedad = Me.Plazo
        Exit Function
    End If
    If m_NombrePropiedad = "Calidad" Then
        getPropiedad = Me.Calidad
        Exit Function
    End If
    If m_NombrePropiedad = "Coste" Then
        getPropiedad = Me.Coste
        Exit Function
    End If
    If m_NombrePropiedad = "ImpactoGlobal" Then
        getPropiedad = Me.ImpactoGlobal
        Exit Function
    End If
    If m_NombrePropiedad = "Vulnerabilidad" Then
        getPropiedad = Me.Vulnerabilidad
        Exit Function
    End If
    If m_NombrePropiedad = "Valoracion" Then
        getPropiedad = Me.Valoracion
        Exit Function
    End If
    If m_NombrePropiedad = "Mitigacion" Then
        getPropiedad = Me.Mitigacion
        Exit Function
    End If
    If m_NombrePropiedad = "Contingencia" Then
        getPropiedad = Me.Contingencia
        Exit Function
    End If
    If m_NombrePropiedad = "RequierePlanContingencia" Then
        getPropiedad = Me.RequierePlanContingencia
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "CausaRaiz" Then
        getPropiedad = Me.CausaRaiz
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        getPropiedad = Me.Estado
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEstado" Then
        getPropiedad = Me.FechaEstado
        Exit Function
    End If
    If m_NombrePropiedad = "Priorizacion" Then
        getPropiedad = Me.Priorizacion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaMaterializado" Then
        getPropiedad = Me.FechaMaterializado
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRetirado" Then
        getPropiedad = Me.FechaRetirado
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCerrado" Then
        getPropiedad = Me.FechaCerrado
        Exit Function
    End If
    If m_NombrePropiedad = "FechaMitigacionAceptar" Then
        getPropiedad = Me.FechaMitigacionAceptar
        Exit Function
    End If
    If m_NombrePropiedad = "JustificacionAceptacionRiesgo" Then
        getPropiedad = Me.JustificacionAceptacionRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "FechaJustificacionAceptacionRiesgo" Then
        getPropiedad = Me.FechaJustificacionAceptacionRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAprobacionAceptacionPorCalidad" Then
        getPropiedad = Me.FechaAprobacionAceptacionPorCalidad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRechazoAceptacionPorCalidad" Then
        getPropiedad = Me.FechaRechazoAceptacionPorCalidad
        Exit Function
    End If
    If m_NombrePropiedad = "JustificacionRetiroRiesgo" Then
        getPropiedad = Me.JustificacionRetiroRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "FechaJustificacionRetiroRiesgo" Then
        getPropiedad = Me.FechaJustificacionRetiroRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAprobacionRetiroPorCalidad" Then
        getPropiedad = Me.FechaAprobacionRetiroPorCalidad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRechazoRetiroPorCalidad" Then
        getPropiedad = Me.FechaRechazoRetiroPorCalidad
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereRiesgoDeBiblioteca" Then
        getPropiedad = Me.RequiereRiesgoDeBiblioteca
        Exit Function
    End If
    If m_NombrePropiedad = "CodRiesgoBiblioteca" Then
        getPropiedad = Me.CodRiesgoBiblioteca
        Exit Function
    End If
    If m_NombrePropiedad = "RiesgoPendienteRetipificacion" Then
        getPropiedad = Me.RiesgoPendienteRetipificacion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRiesgoParaRetipificar" Then
        getPropiedad = Me.FechaRiesgoParaRetipificar
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRiesgoRetipificado" Then
        getPropiedad = Me.FechaRiesgoRetipificado
        Exit Function
    End If
    If m_NombrePropiedad = "DiasSinRespuestaCalidadAceptacion" Then
        getPropiedad = Me.DiasSinRespuestaCalidadAceptacion
        Exit Function
    End If
    If m_NombrePropiedad = "DiasSinRespuestaCalidadRetiro" Then
        getPropiedad = Me.DiasSinRespuestaCalidadRetiro
        Exit Function
    End If
    If m_NombrePropiedad = "DiasSinRespuestaCalidadRetipificacion" Then
        getPropiedad = Me.DiasSinRespuestaCalidadRetipificacion
        Exit Function
    End If
    If m_NombrePropiedad = "Origen" Then
        getPropiedad = Me.Origen
        Exit Function
    End If
    If m_NombrePropiedad = "HayErrorEnRiesgo" Then
        getPropiedad = Me.HayErrorEnRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "NombreNodoDesc" Then
        getPropiedad = Me.NombreNodoDesc
        Exit Function
    End If
    If m_NombrePropiedad = "NombreNodoEstado" Then
        getPropiedad = Me.NombreNodoEstado
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIcono" Then
        getPropiedad = Me.NombreIcono
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Riesgo.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDRiesgo" Then
        Me.IDRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        Me.IDEdicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoUnico" Then
        Me.CodigoUnico = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoRiesgo" Then
        Me.CodigoRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDetectado" Then
        Me.FechaDetectado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DetectadoPor" Then
        Me.DetectadoPor = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EntidadDetecta" Then
        Me.EntidadDetecta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Plazo" Then
        Me.Plazo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Calidad" Then
        Me.Calidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Coste" Then
        Me.Coste = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ImpactoGlobal" Then
        Me.ImpactoGlobal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Vulnerabilidad" Then
        Me.Vulnerabilidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Valoracion" Then
        Me.Valoracion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Mitigacion" Then
        Me.Mitigacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Contingencia" Then
        Me.Contingencia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RequierePlanContingencia" Then
        Me.RequierePlanContingencia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CausaRaiz" Then
        Me.CausaRaiz = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        Me.Estado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEstado" Then
        Me.FechaEstado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Priorizacion" Then
        Me.Priorizacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaMaterializado" Then
        Me.FechaMaterializado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRetirado" Then
        Me.FechaRetirado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCerrado" Then
        Me.FechaCerrado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaMitigacionAceptar" Then
        Me.FechaMitigacionAceptar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "JustificacionAceptacionRiesgo" Then
        Me.JustificacionAceptacionRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaJustificacionAceptacionRiesgo" Then
        Me.FechaJustificacionAceptacionRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAprobacionAceptacionPorCalidad" Then
        Me.FechaAprobacionAceptacionPorCalidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRechazoAceptacionPorCalidad" Then
        Me.FechaRechazoAceptacionPorCalidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "JustificacionRetiroRiesgo" Then
        Me.JustificacionRetiroRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaJustificacionRetiroRiesgo" Then
        Me.FechaJustificacionRetiroRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAprobacionRetiroPorCalidad" Then
        Me.FechaAprobacionRetiroPorCalidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRechazoRetiroPorCalidad" Then
        Me.FechaRechazoRetiroPorCalidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereRiesgoDeBiblioteca" Then
        Me.RequiereRiesgoDeBiblioteca = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodRiesgoBiblioteca" Then
        Me.CodRiesgoBiblioteca = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RiesgoPendienteRetipificacion" Then
        Me.RiesgoPendienteRetipificacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRiesgoParaRetipificar" Then
        Me.FechaRiesgoParaRetipificar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRiesgoRetipificado" Then
        Me.FechaRiesgoRetipificado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DiasSinRespuestaCalidadAceptacion" Then
        Me.DiasSinRespuestaCalidadAceptacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DiasSinRespuestaCalidadRetiro" Then
        Me.DiasSinRespuestaCalidadRetiro = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DiasSinRespuestaCalidadRetipificacion" Then
        Me.DiasSinRespuestaCalidadRetipificacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Origen" Then
        Me.Origen = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HayErrorEnRiesgo" Then
        Me.HayErrorEnRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreNodoDesc" Then
        Me.NombreNodoDesc = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreNodoEstado" Then
        Me.NombreNodoEstado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreIcono" Then
        Me.NombreIcono = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Riesgo.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Riesgo.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function



