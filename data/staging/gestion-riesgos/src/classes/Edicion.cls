VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Event EdicionEliminada(ByVal p_IDEdicion As String, ByVal p_IDProyecto As String)

Public Event NuevaEdicion(ByVal p_ObjEdicion As Edicion)
Public Event PublicacionEliminada()


Public IDEdicion As String
Public IDProyecto As String
Public FechaEdicion As String
Public Edicion As String
Public Elaborado As String
Public Revisado As String
Public Aprobado As String
Public FechaPublicacion As String
Public EntregadoAClienteORAC As String
Public Comentarios As String
Public PermitidoImprimirExcel As String
Public IDDocumentoAGEDO As String
Public NombreArchivoInforme As String
Public FechaMaxProximaPublicacion As String

Public FechaPreparadaParaPublicar As String
Public UsuarioProponePublicar As String
Public PropuestaRechazadaPorCalidadFecha As String
Public PropuestaRechazadaPorCalidadMotivo As String
Public UsuarioCalidadRechazaPropuesta As String
Public NotasCalidadParaPublicar As String
Public FechaUltimoCambio As String
Public UsuarioUltimoCambio As String

Private m_sIDEdicionCalculada As String
Private m_sFechaMaxProximaPublicacionCalculado As String
Private m_sFechaUltimaPublicacion As String
Private m_eEsUltimaEdicionCaltulado As EnumSiNo
Private m_sEsUltimaEdicion As String
Private m_eEsPrimeraEdicion As EnumSiNo
Private m_ObjProyecto As Proyecto
Private m_ObjUsuarioElabora As Usuario
Private m_ObjUsuarioRevisa As Usuario
Private m_ObjUsuarioAprueba As Usuario
Private m_ObjDocumentoObj As Documento
Private m_objColRiesgos As Scripting.Dictionary

Private m_ObjEdicionUltima As Edicion

Private m_ObjEdicionAnterior As Edicion
Private m_ObjEdicionSiguiente As Edicion
Private m_ObjEdicionPrimera As Edicion
Private m_eEsActivo As EnumSiNo
Private m_sCodigoRiesgoCalculado As String
Private m_sCodigoUnicoRiesgoNuevo As String
Private m_eTieneAnexoEvidenciaUTE As EnumSiNo
Private m_sURLArchivoInforme As String
Private m_eOrdenarColeRiesgosAscendentemente As EnumSiNo
Private m_sURLEvidenciaUTE As String
Private m_ObjColCamposRiesgoParaCambios As Collection
Private m_eUsuarioConectadoAutorizado As EnumSiNo
Private m_ObjColAnexos As Scripting.Dictionary
Private m_objColRiesgosPorPrioridad As Scripting.Dictionary
Private m_objColRiesgosPorPrioridadNoRetirados As Scripting.Dictionary
Private m_objColRiesgosPorPrioridadTodos As Scripting.Dictionary

Private m_ObjColRiesgosOferta As Scripting.Dictionary

Private m_objColRiesgosExternos As Scripting.Dictionary

Private m_eTieneRiesgos As EnumSiNo
Private m_eTieneRiesgosExt As EnumSiNo
Private m_eTieneRiesgosDeOferta As EnumSiNo

Private m_objcolCambiosConEdicionAnterior As Scripting.Dictionary
Private m_eAlgunRiesgoNoCerrado As EnumSiNo
Private m_eAlgunRiesgoNoDeBiblioteca As EnumSiNo
Private m_eTodosLosRiesgosPriorizados As EnumSiNo
Private m_eAlgunRiesgoPriorizado As EnumSiNo

Private m_eAlgunCambioConEdicionAnteriorRegistrado As EnumSiNo
Private m_ObjcolRiesgosNoRetirados As Scripting.Dictionary
Private m_ObjcolRiesgosActivos As Scripting.Dictionary
Private m_objColEvolucionRiesgos As Scripting.Dictionary
Private m_sCorreoRAC As String
Private m_sParaInformeAvisos As String

Private m_ePreparadaParaPublicar As EnumSiNo
Private m_sCorreoResponsableCalidad As String
Private m_objUsuarioProponePublicarObj As Usuario
Private m_sNombreUsuarioProponePublicar As String

Private m_objColLogPulicaciones As Scripting.Dictionary
Private m_objColLogPublicacionesProyecto As Scripting.Dictionary
Private m_objColLogCorreosResponsables As Scripting.Dictionary

Private m_eAlgunHIstorialPublicacion As EnumSiNo
Private m_eEdicionPrimeraYActivaConPermisoUsuario As EnumSiNo
Private m_ePublicable As EnumSiNo
Private m_eEvidenciasSuministradoresRequeridas As EnumSiNo
Private m_eEvidenciasSuministradoresCompletadas As EnumSiNo
Private m_sDiasSinPublicar As String
Private m_objColSuministradoresEvidencias As Scripting.Dictionary
Private m_objColSuministradores As Scripting.Dictionary

Private m_sHTMLEdicionQuitarPropuestaPublicacion As String
Private m_sHTMLEdicionRechazarPropuestaPublicacion As String
Private m_sHTMLEdicionPropuestaPublicacion As String
Private m_sHTMLEdicionRevision As String
Private m_sHTMLHistorialPublicaciones As String
Private m_sURLHistorialPublicaciones As String
Private m_sHTMLHistorialcorreosResponsables As String
Private m_objAnexoEvidenciaUTE As Anexo

Private m_ePublicacionRechazada As EnumSiNo
Private m_sNombreParaNodo As String
Private m_sNombreParaNodo1 As String
Private m_ObjUsuarioUltimoCambioObj As Usuario

Private m_ObjColSuministradoresFaltan As Scripting.Dictionary
Private m_ObjColSuministradoresSobran As Scripting.Dictionary

Private m_objColCampos As Collection
Private m_objColCamposParaCambios As Collection
Public Error As String
Private m_Error As String


Public Property Get UsuarioUltimoCambioObj() As Usuario
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuarioUltimoCambioObj Is Nothing Then
        Set UsuarioUltimoCambioObj = m_ObjUsuarioUltimoCambioObj
        Exit Property
    End If
    If Me.UsuarioUltimoCambio = "" Then
        Exit Property
    End If
    Set m_ObjUsuarioUltimoCambioObj = Constructor.getUsuario(p_UsuarioRed:=Me.UsuarioUltimoCambio, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioUltimoCambioObj = m_ObjUsuarioUltimoCambioObj
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.UsuarioUltimoCambioObj ha devuelto el error: " & vbNewLine & Err.Description
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
    With Me
        m_sNombreParaNodo = .Proyecto.NombreParaNodo & " Ed. " & .Edicion & " " & IIf(.EsActivo = EnumSiNo.Sí, "(Activa)", "(No Activa)")
    End With
    
    NombreParaNodo = m_sNombreParaNodo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.NombreParaNodo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NombreParaNodo1() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sNombreParaNodo1 <> "" Then
        NombreParaNodo1 = m_sNombreParaNodo1
        Exit Property
    End If
    m_sNombreParaNodo1 = Me.Proyecto.NombreParaNodo & " (" & UCase(Me.Proyecto.NombreUsuarioCalidadParaNodo) & ")"
  
    
    NombreParaNodo1 = m_sNombreParaNodo1
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.NombreParaNodo1 ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get NombreUsuarioProponePublicar() As String
    
     Me.Error = ""
    m_Error = ""
    
    If Me.UsuarioProponePublicarObj Is Nothing Then
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
    m_sNombreUsuarioProponePublicar = Me.UsuarioProponePublicarObj.Nombre
    NombreUsuarioProponePublicar = m_sNombreUsuarioProponePublicar
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.NombreUsuarioProponePublicar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get UsuarioProponePublicarObj() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objUsuarioProponePublicarObj Is Nothing Then
        Set UsuarioProponePublicarObj = m_objUsuarioProponePublicarObj
        Exit Property
    End If
    If Me.UsuarioProponePublicar = "" Then
        Exit Property
    End If
    Set m_objUsuarioProponePublicarObj = getUsuario(, Me.UsuarioProponePublicar, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioProponePublicarObj = m_objUsuarioProponePublicarObj
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.UsuarioProponePublicarObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get CorreoResponsableCalidad() As String
    
    Dim m_Proyecto As Proyecto
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sCorreoResponsableCalidad <> "" Then
        CorreoResponsableCalidad = m_sCorreoResponsableCalidad
        Exit Property
    End If
    Set m_Proyecto = Me.Proyecto
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Proyecto Is Nothing Then
        Exit Property
    End If
    
    m_sCorreoResponsableCalidad = m_Proyecto.CorreoResponsableCalidad
    CorreoResponsableCalidad = m_sCorreoResponsableCalidad
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.CorreoResponsableCalidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get PreparadaParaPublicar() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
    If Me.EsActivo = EnumSiNo.No Then
        m_ePreparadaParaPublicar = EnumSiNo.No
        PreparadaParaPublicar = m_ePreparadaParaPublicar
        Exit Property
    End If
    If IsDate(Me.FechaPreparadaParaPublicar) Then
        m_ePreparadaParaPublicar = EnumSiNo.Sí
    Else
        m_ePreparadaParaPublicar = EnumSiNo.No
    End If
    
    PreparadaParaPublicar = m_ePreparadaParaPublicar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.PreparadaParaPublicar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get AlgunHIstorialPublicacion() As EnumSiNo
    
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""

    If m_eAlgunHIstorialPublicacion <> Empty Then
        AlgunHIstorialPublicacion = m_eAlgunHIstorialPublicacion
        Exit Property
    End If
    Set m_Col = Me.ColLogPublicacionesProyecto
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_eAlgunHIstorialPublicacion = EnumSiNo.No
    Else
        m_eAlgunHIstorialPublicacion = EnumSiNo.Sí
    End If
    
    
    AlgunHIstorialPublicacion = m_eAlgunHIstorialPublicacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.AlgunHIstorialPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ParaInformeAvisos() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Proyecto Is Nothing Then
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
    m_sParaInformeAvisos = Me.Proyecto.ParaInformeAvisos
    ParaInformeAvisos = m_sParaInformeAvisos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ParaInformeAvisos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasSinPublicar() As String
    
    Dim m_Fecha As String
    Dim m_FechaRef As Date
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    If Not IsDate(Me.FechaPreparadaParaPublicar) Then
        Exit Property
    End If
    If IsDate(Me.FechaPublicacion) Then
        m_Fecha = Me.FechaPublicacion
    Else
        m_Fecha = CStr(m_FechaRef)
    End If
    m_sDiasSinPublicar = DateDiff("d", CDate(Me.FechaPreparadaParaPublicar), CDate(m_Fecha))
    
    DiasSinPublicar = m_sDiasSinPublicar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.DiasSinPublicar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HTMLEdicionQuitarPropuestaPublicacion() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sHTMLEdicionQuitarPropuestaPublicacion <> "" Then
        HTMLEdicionQuitarPropuestaPublicacion = m_sHTMLEdicionQuitarPropuestaPublicacion
        Exit Property
    End If
    
    m_sHTMLEdicionQuitarPropuestaPublicacion = getHTMLEdicionQuitarPropuestaPublicacion(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    HTMLEdicionQuitarPropuestaPublicacion = m_sHTMLEdicionQuitarPropuestaPublicacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.HTMLEdicionQuitarPropuestaPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HTMLEdicionRechazarPropuestaPublicacion() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sHTMLEdicionRechazarPropuestaPublicacion <> "" Then
        HTMLEdicionRechazarPropuestaPublicacion = m_sHTMLEdicionRechazarPropuestaPublicacion
        Exit Property
    End If
    
    m_sHTMLEdicionRechazarPropuestaPublicacion = getHTMLEdicionRechazarPropuestaPublicacion(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    HTMLEdicionRechazarPropuestaPublicacion = m_sHTMLEdicionRechazarPropuestaPublicacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.HTMLEdicionRechazarPropuestaPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HTMLEdicionPropuestaPublicacion() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sHTMLEdicionPropuestaPublicacion <> "" Then
        HTMLEdicionPropuestaPublicacion = m_sHTMLEdicionPropuestaPublicacion
        Exit Property
    End If
    
    m_sHTMLEdicionPropuestaPublicacion = getHTMLEdicionPropuestaPublicacion(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    HTMLEdicionPropuestaPublicacion = m_sHTMLEdicionPropuestaPublicacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.HTMLEdicionPropuestaPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get HTMLEdicionRevision() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sHTMLEdicionRevision <> "" Then
        HTMLEdicionRevision = m_sHTMLEdicionRevision
        Exit Property
    End If
    
    m_sHTMLEdicionRevision = getHTMLEdicionRevision(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    HTMLEdicionRevision = m_sHTMLEdicionRevision
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.HTMLEdicionRevision ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get HTMLHistorialPublicaciones() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
   
    
    m_sHTMLHistorialPublicaciones = getHTMLHistorialPublicaciones(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    HTMLHistorialPublicaciones = m_sHTMLHistorialPublicaciones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.HTMLHistorialPublicaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLHistorialPublicaciones() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sURLHistorialPublicaciones <> "" Then
        URLHistorialPublicaciones = m_sURLHistorialPublicaciones
        Exit Property
    End If
    
    m_sURLHistorialPublicaciones = getURLHistorialPublicaciones(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    URLHistorialPublicaciones = m_sURLHistorialPublicaciones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.URLHistorialPublicaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get HTMLHistorialcorreosResponsables() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
   
    
    m_sHTMLHistorialcorreosResponsables = getHTMLHistorialcorreosResponsables(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    HTMLHistorialcorreosResponsables = m_sHTMLHistorialcorreosResponsables
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.HTMLHistorialcorreosResponsables ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Private Function getURLHistorialPublicaciones( _
                                            Optional ByRef p_Error As String _
                                            ) As String

   
    Dim m_HTMLHistorialPublicaciones As String
    
    
    On Error GoTo errores
    
    If Me.ColLogPublicacionesProyecto Is Nothing Then
        Exit Function
    End If
    m_HTMLHistorialPublicaciones = Me.HTMLHistorialPublicaciones
    p_Error = Me.Error
    If p_Error <> "" Then
        p_Error = "Nada registrado"
        Err.Raise 1000
    End If
    
    
    getURLHistorialPublicaciones = HTMLENTXT(m_HTMLHistorialPublicaciones, , EnumSiNo.No, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getURLHistorialPublicaciones ha producido el error num " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If

End Function
Public Property Get CorreoRAC() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.Proyecto Is Nothing Then
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
    m_sCorreoRAC = Me.Proyecto.CorreoRAC
    CorreoRAC = m_sCorreoRAC
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.CorreoRAC ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get OrdenarColeRiesgosAscendentemente() As EnumSiNo
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eOrdenarColeRiesgosAscendentemente <> 0 Then
        OrdenarColeRiesgosAscendentemente = m_eOrdenarColeRiesgosAscendentemente
        Exit Property
    End If
    m_eOrdenarColeRiesgosAscendentemente = EnumSiNo.Sí
    OrdenarColeRiesgosAscendentemente = m_eOrdenarColeRiesgosAscendentemente
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método OrdenarColeRiesgosAscendentemente ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let OrdenarColeRiesgosAscendentemente(ByVal eNewValue As EnumSiNo)

    m_eOrdenarColeRiesgosAscendentemente = eNewValue

End Property
Public Property Get FechaMaxProximaPublicacionCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
'    If IsDate(m_sFechaMaxProximaPublicacionCalculado) Then
'        m_sFechaMaxProximaPublicacionCalculado = Format(m_sFechaMaxProximaPublicacionCalculado, "dd/mm/yyyy")
'        FechaMaxProximaPublicacionCalculado = m_sFechaMaxProximaPublicacionCalculado
'        Exit Property
'    End If
    If Me.Proyecto Is Nothing Then
        Exit Property
    End If
    m_sFechaMaxProximaPublicacionCalculado = Me.Proyecto.FechaMaxProximaPublicacionCalculada
    m_Error = Me.Proyecto.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    FechaMaxProximaPublicacionCalculado = m_sFechaMaxProximaPublicacionCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.FechaMaxProximaPublicacionCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get FechaUltimaPublicacion() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If IsDate(m_sFechaUltimaPublicacion) Then
        m_sFechaUltimaPublicacion = Format(m_sFechaUltimaPublicacion, "dd/mm/yyyy")
        FechaUltimaPublicacion = m_sFechaUltimaPublicacion
        Exit Property
    End If
    If Me.Proyecto Is Nothing Then
        Exit Property
    End If
    m_sFechaUltimaPublicacion = Me.Proyecto.FechaUltimaPublicacion
    m_Error = Me.Proyecto.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    FechaUltimaPublicacion = m_sFechaUltimaPublicacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.FechaUltimaPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get Proyecto() As Proyecto
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjProyecto Is Nothing Then
        Set Proyecto = m_ObjProyecto
        Exit Property
    End If
    If Me.IDProyecto = "" Then
        m_Error = "No se conoce el IDProyecto"
        Err.Raise 1000
    End If
    Set m_ObjProyecto = Constructor.getProyecto(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Proyecto = m_ObjProyecto
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.Proyecto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property
Public Property Set Proyecto(ByVal sNewValue As Variant)
    
    Set m_ObjProyecto = sNewValue

End Property
Public Property Get UsuarioElabora() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
     m_Error = ""
    If Not m_ObjUsuarioElabora Is Nothing Then
        Set UsuarioElabora = m_ObjUsuarioElabora
        Exit Property
    End If
    If Me.Elaborado = "" Then
        m_Error = "No se conoce el Elaborado"
        Err.Raise 1000
    End If
    Set m_ObjUsuarioElabora = getUsuario(, , Me.Elaborado, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioElabora = m_ObjUsuarioElabora
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.UsuarioElabora ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property

Public Property Set UsuarioElabora(ByVal sNewValue As Variant)
    
   
    Set m_ObjUsuarioElabora = sNewValue

End Property

Public Property Get UsuarioRevisa() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
     m_Error = ""
    If Not m_ObjUsuarioRevisa Is Nothing Then
        Set UsuarioRevisa = m_ObjUsuarioRevisa
        Exit Property
    End If
    If Me.Revisado = "" Then
        m_Error = "No se conoce el Revisado"
        Err.Raise 1000
    End If
    Set m_ObjUsuarioRevisa = getUsuario(, , Me.Revisado, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioRevisa = m_ObjUsuarioRevisa
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.UsuarioRevisa ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set UsuarioRevisa(ByVal sNewValue As Variant)
    
    
    Set m_ObjUsuarioRevisa = sNewValue

End Property

Public Property Get UsuarioAprueba() As Usuario

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuarioAprueba Is Nothing Then
        Set UsuarioAprueba = m_ObjUsuarioAprueba
        Exit Property
    End If
    If Me.Aprobado = "" Then
        m_Error = "No se conoce el Aprobado"
        Err.Raise 1000
    End If
    Set m_ObjUsuarioAprueba = getUsuario(, , Me.Aprobado, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioAprueba = m_ObjUsuarioAprueba
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.UsuarioAprueba ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set UsuarioAprueba(ByVal sNewValue As Variant)
    
    
    Set m_ObjUsuarioAprueba = sNewValue

End Property

Public Property Get DocumentoObj() As Documento
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjDocumentoObj Is Nothing Then
        Set DocumentoObj = m_ObjDocumentoObj
        Exit Property
    End If
    If Me.IDDocumentoAGEDO = "" Then
        Exit Property
    End If
    Set m_ObjDocumentoObj = getDocumentoByID(Me.IDDocumentoAGEDO, m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set DocumentoObj = m_ObjDocumentoObj
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.DocumentoObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set DocumentoObj(ByVal sNewValue As Variant)
    
    Set m_ObjDocumentoObj = sNewValue

End Property

Public Property Get colRiesgos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgos Is Nothing Then
        Set colRiesgos = m_objColRiesgos
        Exit Property
    End If
    Set m_objColRiesgos = Constructor.getRiesgosPorEdicion(Me.IDEdicion, EnumSiNo.No, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colRiesgos = m_objColRiesgos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.colRiesgos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colRiesgos(ByVal sNewValue As Variant)
    
    Set m_objColRiesgos = sNewValue
    Set m_objColRiesgosPorPrioridad = Nothing

End Property

Public Property Get colRiesgosNoRetirados() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjcolRiesgosNoRetirados Is Nothing Then
        Set colRiesgosNoRetirados = m_ObjcolRiesgosNoRetirados
        Exit Property
    End If
    Set m_ObjcolRiesgosNoRetirados = Constructor.getRiesgosNoRetiradosNoOrdenados(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colRiesgosNoRetirados = m_ObjcolRiesgosNoRetirados
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.colRiesgosNoRetirados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colRiesgosNoRetirados(ByVal sNewValue As Variant)
    
    Set m_ObjcolRiesgosNoRetirados = sNewValue
    Set m_objColRiesgosPorPrioridadNoRetirados = Nothing
    Set m_ObjcolRiesgosActivos = Nothing
End Property

Public Property Get colRiesgosActivos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjcolRiesgosActivos Is Nothing Then
        Set colRiesgosActivos = m_ObjcolRiesgosActivos
        Exit Property
    End If
    Set m_ObjcolRiesgosActivos = Constructor.getRiesgosPorEdicionActivos(Me.IDEdicion, EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colRiesgosActivos = m_ObjcolRiesgosActivos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.colRiesgosActivos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colRiesgosActivos(ByVal sNewValue As Variant)
    
    Set m_ObjcolRiesgosActivos = sNewValue
    

End Property
Public Property Get ColRiesgosPorPrioridadNoRetirados() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosPorPrioridadNoRetirados Is Nothing Then
        Set ColRiesgosPorPrioridadNoRetirados = m_objColRiesgosPorPrioridadNoRetirados
        Exit Property
    End If
    If Me.TodosLosRiesgosPriorizados = EnumSiNo.No Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    Set m_objColRiesgosPorPrioridadNoRetirados = Constructor.getRiesgosPorPrioridad(Me.IDEdicion, EnumSiNo.No, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRiesgosPorPrioridadNoRetirados = m_objColRiesgosPorPrioridadNoRetirados
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColRiesgosPorPrioridadNoRetirados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set ColRiesgosPorPrioridadNoRetirados(ByVal sNewValue As Variant)
    
    Set m_objColRiesgosPorPrioridadNoRetirados = Nothing
    
    
End Property
Public Property Get ColRiesgosPorPrioridadTodos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosPorPrioridadTodos Is Nothing Then
        Set ColRiesgosPorPrioridadTodos = m_objColRiesgosPorPrioridadTodos
        Exit Property
    End If
    If Me.TodosLosRiesgosPriorizados = EnumSiNo.No Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    Set m_objColRiesgosPorPrioridadTodos = Constructor.getRiesgosPorPrioridadTodos(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRiesgosPorPrioridadTodos = m_objColRiesgosPorPrioridadTodos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.ColRiesgosPorPrioridadTodos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set ColRiesgosPorPrioridadTodos(ByVal sNewValue As Variant)
    
    Set m_objColRiesgosPorPrioridadTodos = Nothing
    
    
End Property

Public Property Get IDEdicionCalculada() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sIDEdicionCalculada = DameID("TbProyectosEdiciones", "IDEdicion", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDEdicionCalculada = m_sIDEdicionCalculada
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.IDEdicionCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get AnexoEvidenciaUTE() As Anexo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objAnexoEvidenciaUTE Is Nothing Then
        Set AnexoEvidenciaUTE = m_objAnexoEvidenciaUTE
        Exit Property
    End If
    Set m_objAnexoEvidenciaUTE = Constructor.getAnexoEvidenciaUTE(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set AnexoEvidenciaUTE = m_objAnexoEvidenciaUTE
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.AnexoEvidenciaUTE ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set AnexoEvidenciaUTE(ByVal objNewValue As Anexo)

    Set m_objAnexoEvidenciaUTE = objNewValue

End Property

Public Property Get TieneAnexoEvidenciaUTE() As EnumSiNo
    
    Dim m_ObjAnexo As Anexo
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjAnexo = Me.AnexoEvidenciaUTE
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjAnexo Is Nothing Then
        m_eTieneAnexoEvidenciaUTE = EnumSiNo.No
    Else
        m_eTieneAnexoEvidenciaUTE = EnumSiNo.Sí
    End If
    
    TieneAnexoEvidenciaUTE = m_eTieneAnexoEvidenciaUTE
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.TieneAnexoEvidenciaUTE ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let TieneAnexoEvidenciaUTE(ByVal eNewValue As EnumSiNo)
   
    m_eTieneAnexoEvidenciaUTE = eNewValue

End Property

Public Property Get EsUltimaEdicionCalculado() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.IDEdicion = "" Then
        m_Error = "El método Edicion.EsUltimaEdicionCalculado ha llamado al método Edicion.IDEdicion sin resultado"
        Err.Raise 1000
    End If
    If Me.EdicionUltima Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido determinar cuál es la última Edición del Proyecto"
        Err.Raise 1000
    End If
    If Me.EdicionUltima.IDEdicion = Me.IDEdicion Then
        m_eEsUltimaEdicionCaltulado = EnumSiNo.Sí
    Else
        m_eEsUltimaEdicionCaltulado = EnumSiNo.No
    End If
    EsUltimaEdicionCalculado = m_eEsUltimaEdicionCaltulado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.EsUltimaEdicionCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property

Public Property Let EsUltimaEdicionCalculado(ByVal eNewValue As EnumSiNo)
   
    m_eEsUltimaEdicionCaltulado = eNewValue

End Property
Public Property Get EsPrimeraEdicion() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.IDEdicion = "" Then
        m_Error = "El método Edicion.EsPrimeraEdicion ha llamado al método Edicion.IDEdicion sin resultado"
        Err.Raise 1000
    End If
    If Me.EdicionAnterior Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eEsPrimeraEdicion = EnumSiNo.Sí
    Else
        m_eEsPrimeraEdicion = EnumSiNo.No
    End If
    EsPrimeraEdicion = m_eEsPrimeraEdicion
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.EsPrimeraEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property

Public Property Let EsPrimeraEdicion(ByVal eNewValue As EnumSiNo)
   
    m_eEsPrimeraEdicion = eNewValue

End Property

Public Property Get EdicionUltima() As Edicion
    
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.IDProyecto = "" Then
        m_Error = "El método Edicion.getEdicionUltima ha llamado al método EdicionPrimera.IDProyecto sin resultado"
        Err.Raise 1000
    End If
    Set m_ObjEdicionUltima = Constructor.getEdicionUltima(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EdicionUltima = m_ObjEdicionUltima
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.EdicionUltima ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set EdicionUltima(ByVal objNewValue As Variant)
   
    Set m_ObjEdicionUltima = objNewValue

End Property

Public Property Get EdicionAnterior() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_ObjEdicionAnterior = Constructor.getEdicionAnterior(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set EdicionAnterior = m_ObjEdicionAnterior
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.EdicionAnterior ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set EdicionAnterior(ByVal objNewValue As Variant)
    
    Set m_ObjEdicionAnterior = objNewValue

End Property

Public Property Get EdicionSiguiente() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
     m_Error = ""
    If Not m_ObjEdicionSiguiente Is Nothing Then
        Set EdicionSiguiente = m_ObjEdicionSiguiente
        Exit Property
    End If
    
    Set m_ObjEdicionSiguiente = Constructor.getEdicionSiguiente(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EdicionSiguiente = m_ObjEdicionSiguiente
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.EdicionSiguiente ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set EdicionSiguiente(ByVal objNewValue As Variant)
   
    Set m_ObjEdicionSiguiente = objNewValue

End Property

Public Property Get EdicionPrimera() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEdicionPrimera Is Nothing Then
        Set EdicionPrimera = m_ObjEdicionPrimera
        Exit Property
    End If
    Set m_ObjEdicionPrimera = Constructor.getEdicionPrimera(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EdicionPrimera = m_ObjEdicionPrimera
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.EdicionPrimera ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set EdicionPrimera(ByVal objNewValue As Variant)
    
    Set m_ObjEdicionPrimera = objNewValue

End Property


Public Property Get EsActivo() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eEsActivo <> 0 Then
        EsActivo = m_eEsActivo
        Exit Property
    End If
    If Me.IDEdicion = "" Then
        m_Error = "No se puede determinar la IDEdicion"
        Err.Raise 1000
    End If
    If Me.EdicionUltima Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se puede determinar la última Edición"
        Err.Raise 1000
    End If
    If Me.EdicionUltima.IDEdicion = Me.IDEdicion Then
        If IsDate(Me.FechaPublicacion) Then
            m_eEsActivo = EnumSiNo.No
        Else
            m_eEsActivo = EnumSiNo.Sí
        End If
    Else
        m_eEsActivo = EnumSiNo.No
    End If
    EsActivo = m_eEsActivo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.EsActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Let EsActivo(p_Valor As EnumSiNo)

    m_eEsActivo = p_Valor
End Property
Public Property Get CodigoRiesgoCalculado() As String
    
    Dim m_MaxOrdinalCodigoRiesgo As Integer
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sCodigoRiesgoCalculado <> "" Then
        CodigoRiesgoCalculado = m_sCodigoRiesgoCalculado
        Exit Property
    End If
    If Not Me.colRiesgos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_IdRiesgo In Me.colRiesgos.keys
            Set m_ObjRiesgo = Me.colRiesgos(m_IdRiesgo)
            If IsNumeric(Right(m_ObjRiesgo.CodigoRiesgo, 3)) Then
                If CInt(Right(m_ObjRiesgo.CodigoRiesgo, 3)) > m_MaxOrdinalCodigoRiesgo Then
                    m_MaxOrdinalCodigoRiesgo = CInt(Right(m_ObjRiesgo.CodigoRiesgo, 3))
                End If
            End If
            Set m_ObjRiesgo = Nothing
        Next
        
    End If
    m_sCodigoRiesgoCalculado = "R" & Format(m_MaxOrdinalCodigoRiesgo + 1, "000")
    CodigoRiesgoCalculado = m_sCodigoRiesgoCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.CodigoRiesgoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get CodigoUnicoRiesgoNuevo() As String
    
    Dim m_CodigoRiesgoCalculado As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sCodigoUnicoRiesgoNuevo <> "" Then
        CodigoUnicoRiesgoNuevo = m_sCodigoUnicoRiesgoNuevo
        Exit Property
    End If
    m_CodigoRiesgoCalculado = Me.CodigoRiesgoCalculado
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.Proyecto Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se puede determinar la gestión de riesgos de la Edición"
        Err.Raise 1000
    End If
    m_sCodigoUnicoRiesgoNuevo = Format(Me.IDProyecto, "000") & m_CodigoRiesgoCalculado
    CodigoUnicoRiesgoNuevo = m_sCodigoUnicoRiesgoNuevo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.CodigoUnicoRiesgoNuevo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLArchivoInforme() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLArchivoInforme <> "" Then
        URLArchivoInforme = m_sURLArchivoInforme
        Exit Property
    End If
    If Me.NombreArchivoInforme = "" And Me.IDDocumentoAGEDO = "" Then
        Exit Property
    End If
    If Me.IDDocumentoAGEDO <> "" Then
        If Me.DocumentoObj Is Nothing Then
            m_Error = Me.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            Exit Property
            
        End If
        m_sURLArchivoInforme = Me.DocumentoObj.URLAdjunto
    Else
        If Not fso.FolderExists(m_ObjEntorno.URLDirectorioDocumentacion) Then
            m_Error = "No es alcanzable el directorio de la documentación"
            Err.Raise 1000
        End If
        m_sURLArchivoInforme = m_ObjEntorno.URLDirectorioDocumentacion & Me.NombreArchivoInforme
    End If
    URLArchivoInforme = m_sURLArchivoInforme
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.URLArchivoInforme ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get URLEvidenciaUTE() As String
    
    Dim m_ObjAnexo As Anexo
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set m_ObjAnexo = Me.AnexoEvidenciaUTE
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjAnexo Is Nothing Then
        m_sURLEvidenciaUTE = m_ObjAnexo.URLAnexo
    End If
    
    URLEvidenciaUTE = m_sURLEvidenciaUTE
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.URLEvidenciaUTE ha devuelto el error: " & vbNewLine & Err.Description
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
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.UsuarioConectadoAutorizado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get ColAnexos() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Me.IDEdicion = "" Then
        m_Error = "No se conoce el IDEdicion"
        Err.Raise 1000
    End If
    Set m_ObjColAnexos = Constructor.getAnexosDeEdicion(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnexos = m_ObjColAnexos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColAnexos ha devuelto el error: " & vbNewLine & Err.Description
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
    p_Error = ""

    If Me.IDEdicion = "" Then
        m_Error = "No se conoce el IDEdicion"
        Err.Raise 1000
    End If

    Set ColAnexosTotales = modFormRiesgoDocumentosHelper.getAnexosDeEdicionCached(Me.IDEdicion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Edicion.ColAnexosTotales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = p_Error
End Function


Public Property Get ColRiesgosOferta() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColRiesgosOferta Is Nothing Then
        Set ColRiesgosOferta = m_ObjColRiesgosOferta
        Exit Property
    End If
    Set m_ObjColRiesgosOferta = Constructor.getRiesgosExternos(Me.IDProyecto, EnumOrigenRiesgoExterno.Oferta, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRiesgosOferta = m_ObjColRiesgosOferta
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColRiesgosOferta ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosOferta(ByVal objNewValue As Variant)
    
    Set m_ObjColRiesgosOferta = objNewValue
    
End Property

Public Property Get PublicacionRechazada() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Me.FechaPublicacion <> "" Then
        m_ePublicacionRechazada = EnumSiNo.No
        PublicacionRechazada = m_ePublicacionRechazada
        Exit Property
    End If
    If Me.PropuestaRechazadaPorCalidadFecha = "" Then
        m_ePublicacionRechazada = EnumSiNo.No
    Else
        m_ePublicacionRechazada = EnumSiNo.Sí
    End If
    PublicacionRechazada = m_ePublicacionRechazada
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.PublicacionRechazada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get ColRiesgosExternos() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRiesgosExternos Is Nothing Then
        Set ColRiesgosExternos = m_objColRiesgosExternos
        Exit Property
    End If
    Set m_objColRiesgosExternos = Constructor.getRiesgosExternos(Me.IDProyecto, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRiesgosExternos = m_objColRiesgosExternos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColRiesgosExternos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRiesgosExternos(ByVal objNewValue As Variant)
    
    Set m_objColRiesgosExternos = objNewValue
    If objNewValue Is Nothing Then
        Set Me.ColRiesgosOferta = Nothing
        
    End If
End Property


Public Property Get colCambiosConEdicionAnterior() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objcolCambiosConEdicionAnterior Is Nothing Then
        Set colCambiosConEdicionAnterior = m_objcolCambiosConEdicionAnterior
        Exit Property
    End If
    Set m_objcolCambiosConEdicionAnterior = Constructor.getCambiosPorEdicion(Me.IDProyecto, Me.Edicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colCambiosConEdicionAnterior = m_objcolCambiosConEdicionAnterior
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.colCambiosConEdicionAnterior ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set colCambiosConEdicionAnterior(ByVal objNewValue As Variant)
    
    Set m_objcolCambiosConEdicionAnterior = objNewValue
    
End Property


Public Property Get TieneRiesgos() As EnumSiNo
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Me.colRiesgos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgos = EnumSiNo.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgos = EnumSiNo.Sí
    End If
    TieneRiesgos = m_eTieneRiesgos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.TieneRiesgos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get TieneRiesgosExt() As EnumSiNo
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Me.ColRiesgosExternos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosExt = EnumSiNo.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosExt = EnumSiNo.Sí
    End If
    TieneRiesgosExt = m_eTieneRiesgosExt
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.TieneRiesgosExt ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get TieneRiesgosDeOferta() As EnumSiNo
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If Me.ColRiesgosOferta Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosDeOferta = EnumSiNo.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneRiesgosDeOferta = EnumSiNo.Sí
    End If
    TieneRiesgosDeOferta = m_eTieneRiesgosDeOferta
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.TieneRiesgosDeOferta ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get TodosLosRiesgosPriorizados() As EnumSiNo
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set m_Col = Me.colRiesgosActivos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Property
    End If
    
    For Each m_ID In m_Col
        Set m_ObjRiesgo = m_Col(m_ID)
        If Not IsNumeric(m_ObjRiesgo.Priorizacion) Then
            Set m_ObjRiesgo = Nothing
            m_eTodosLosRiesgosPriorizados = EnumSiNo.No
            TodosLosRiesgosPriorizados = m_eTodosLosRiesgosPriorizados
            Exit Property
        End If
        Set m_ObjRiesgo = Nothing
    Next
        
    m_eTodosLosRiesgosPriorizados = EnumSiNo.Sí
    TodosLosRiesgosPriorizados = m_eTodosLosRiesgosPriorizados
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.TodosLosRiesgosPriorizados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let TodosLosRiesgosPriorizados(ByVal eNewValue As EnumSiNo)

    m_eTodosLosRiesgosPriorizados = eNewValue

End Property

Public Property Get AlgunRiesgoPriorizado() As EnumSiNo
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjRiesgo As riesgo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set m_Col = Me.colRiesgosActivos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Property
    End If
    
    For Each m_ID In m_Col
        Set m_ObjRiesgo = m_Col(m_ID)
        If IsNumeric(m_ObjRiesgo.Priorizacion) Then
            Set m_ObjRiesgo = Nothing
            m_eAlgunRiesgoPriorizado = EnumSiNo.Sí
            AlgunRiesgoPriorizado = m_eAlgunRiesgoPriorizado
            Exit Property
        End If
        Set m_ObjRiesgo = Nothing
    Next
        
    m_eAlgunRiesgoPriorizado = EnumSiNo.No
    AlgunRiesgoPriorizado = m_eAlgunRiesgoPriorizado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.AlgunRiesgoPriorizado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let AlgunRiesgoPriorizado(ByVal eNewValue As EnumSiNo)

    m_eAlgunRiesgoPriorizado = eNewValue

End Property

Public Property Get AlgunCambioConEdicionAnteriorRegistrado() As EnumSiNo
    
    Dim m_Col As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set m_Col = Me.colCambiosConEdicionAnterior
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_eAlgunCambioConEdicionAnteriorRegistrado = EnumSiNo.No
        
    Else
        m_eAlgunCambioConEdicionAnteriorRegistrado = EnumSiNo.Sí
    End If
    
    
        
    
    AlgunCambioConEdicionAnteriorRegistrado = m_eAlgunCambioConEdicionAnteriorRegistrado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.AlgunCambioConEdicionAnteriorRegistrado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let AlgunCambioConEdicionAnteriorRegistrado(ByVal eNewValue As EnumSiNo)

    m_eAlgunCambioConEdicionAnteriorRegistrado = eNewValue

End Property




Private Function getHTMLTablaHistoriaPublicacion( _
                                                    Optional ByRef p_Error As String _
                                                ) As String

    Dim m_mensaje As String
    Dim m_Cabecera As String
    Dim m_ID As Variant
    Dim m_Log As PublicacionLog
    Dim m_Col As Scripting.Dictionary
    
    Const m_Tabulador = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    
    On Error GoTo errores
    
    If Me Is Nothing Then
        p_Error = "No hay datos suficientes"
        Err.Raise 1000
    End If
    Set m_Col = Me.ColLogPublicacionesProyecto
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        getHTMLTablaHistoriaPublicacion = "<p class='tf-empty'>No hay publicaciones registradas para este proyecto.</p>"
        Exit Function
    End If
    If m_Col.Count = 0 Then
        getHTMLTablaHistoriaPublicacion = "<p class='tf-empty'>No hay publicaciones registradas para este proyecto.</p>"
        Exit Function
    End If
    m_mensaje = m_mensaje & "<div class='tf-table-wrap'><table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""ColespanArriba"" colspan='5'>Historial </td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">F.Para Publicar</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">F.Para Publicar Eliminada</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">F. Publicación</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">F Rechazo Propuesta por Calidad</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">Motivo Rechazo Propuesta Por Calidad</td>" & vbNewLine
            
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        For Each m_ID In m_Col
            Set m_Log = m_Col(m_ID)
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                With m_Log
                    If IsDate(.FechaPreparadaParaPublicar) Then
                        m_mensaje = m_mensaje & "<td> " & Format(.FechaPreparadaParaPublicar, "dd/mm/yyyy") & " </td>" & vbNewLine
                    Else
                        m_mensaje = m_mensaje & "<td> &nbsp; </td>" & vbNewLine
                    End If
                    If IsDate(.FechaPreparadaParaPublicarQuitar) Then
                        m_mensaje = m_mensaje & "<td> " & Format(.FechaPreparadaParaPublicarQuitar, "dd/mm/yyyy") & " </td>" & vbNewLine
                    Else
                        m_mensaje = m_mensaje & "<td> &nbsp; </td>" & vbNewLine
                    End If
                    If IsDate(.FechaPublicacion) Then
                        m_mensaje = m_mensaje & "<td> " & Format(.FechaPublicacion, "dd/mm/yyyy") & " </td>" & vbNewLine
                    Else
                        m_mensaje = m_mensaje & "<td> &nbsp; </td>" & vbNewLine
                    End If
                    If IsDate(.PropuestaRechazadaPorCalidadFecha) Then
                        m_mensaje = m_mensaje & "<td> " & Format(.PropuestaRechazadaPorCalidadFecha, "dd/mm/yyyy") & " </td>" & vbNewLine
                    Else
                        m_mensaje = m_mensaje & "<td> &nbsp; </td>" & vbNewLine
                    End If
                    If .PropuestaRechazadaPorCalidadMotivo <> "" Then
                         m_mensaje = m_mensaje & "<td> " & HTMLSafeLargo(.PropuestaRechazadaPorCalidadMotivo) & " </td>" & vbNewLine
                    Else
                        m_mensaje = m_mensaje & "<td> &nbsp; </td>" & vbNewLine
                    End If
                End With
               
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            
            Set m_Log = Nothing
        Next
    m_mensaje = m_mensaje & "</table></div>" & vbNewLine
    getHTMLTablaHistoriaPublicacion = m_mensaje


    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLTablaHistoriaPublicacion ha producido el error num " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If

End Function

Private Function getHTMLTablaHistorialcorreosResponsables( _
                                                            Optional ByRef p_Error As String _
                                                        ) As String

    Dim m_mensaje As String
    Dim m_Cabecera As String
    Dim m_ID As Variant
    Dim m_EdicionCorreoRevision As EdicionCorreoRevision
    Dim m_Col As Scripting.Dictionary
    
    Const m_Tabulador = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    
    On Error GoTo errores
    
    If Me Is Nothing Then
        p_Error = "No hay datos suficientes"
        Err.Raise 1000
    End If
    Set m_Col = Me.ColLogCorreosResponsables
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        getHTMLTablaHistorialcorreosResponsables = "<p class='tf-empty'>No hay correos enviados a responsables para este proyecto.</p>"
        Exit Function
    End If
    If m_Col.Count = 0 Then
        getHTMLTablaHistorialcorreosResponsables = "<p class='tf-empty'>No hay correos enviados a responsables para este proyecto.</p>"
        Exit Function
    End If
    m_mensaje = m_mensaje & "<div class='tf-table-wrap'><table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""ColespanArriba"" colspan='2'>Historial </td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">F.Envío Correo</td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"">IDCorreo</td>" & vbNewLine
           
            
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        For Each m_ID In m_Col
            Set m_EdicionCorreoRevision = m_Col(m_ID)
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                With m_EdicionCorreoRevision
                    If IsDate(.FechaCorreoRevision) Then
                        m_mensaje = m_mensaje & "<td> " & .FechaCorreoRevision & " </td>" & vbNewLine
                    Else
                        m_mensaje = m_mensaje & "<td> &nbsp; </td>" & vbNewLine
                    End If
                    If IsNumeric(.IDCorreo) Then
                        m_mensaje = m_mensaje & "<td> " & .IDCorreo & " </td>" & vbNewLine
                    Else
                        m_mensaje = m_mensaje & "<td> &nbsp; </td>" & vbNewLine
                    End If
                    
                End With
               
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
            
            Set m_EdicionCorreoRevision = Nothing
        Next
    m_mensaje = m_mensaje & "</table></div>" & vbNewLine
    getHTMLTablaHistorialcorreosResponsables = m_mensaje


    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLTablaHistorialcorreosResponsables ha producido el error num " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If

End Function
Private Function getHTMLHistorialPublicaciones( _
                                            Optional ByRef p_Error As String _
                                            ) As String
                                                
    Dim m_HTMLCabecera As String
    Dim m_HTMLTablaHistorial As String
    
    
    Dim m_mensaje As String
    
    On Error GoTo errores
    
    m_HTMLTablaHistorial = getHTMLTablaHistoriaPublicacion(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_mensaje = "<h2 class='tf-section-title'>Historial de publicaciones</h2>" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLTablaHistorial & vbNewLine
    m_mensaje = HTMLDocumentoHistorialTelefonica( _
                    "Historial de publicaciones", _
                    Me.Proyecto.NombreProyecto & " · Edición " & Me.Edicion, _
                    m_mensaje)
    
    getHTMLHistorialPublicaciones = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLHistorialPublicaciones ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getHTMLHistorialcorreosResponsables( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
                                                
    Dim m_HTMLCabecera As String
    Dim m_HTMLTablaHistorial As String
    
    
    Dim m_mensaje As String
    
    On Error GoTo errores
    
    m_HTMLTablaHistorial = getHTMLTablaHistorialcorreosResponsables(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_mensaje = "<h2 class='tf-section-title'>Historial de correos a responsables</h2>" & vbNewLine
    m_mensaje = m_mensaje & m_HTMLTablaHistorial & vbNewLine
    m_mensaje = HTMLDocumentoHistorialTelefonica( _
                    "Historial de correos a responsables", _
                    Me.Proyecto.NombreProyecto & " · Edición " & Me.Edicion, _
                    m_mensaje)
    
    getHTMLHistorialcorreosResponsables = m_mensaje
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLHistorialcorreosResponsables ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getHTMLEdicionPropuestaPublicacion( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
                                                
    Dim m_HTMLTablaProyectoDatos As String
    Dim m_HTMLCuerpo As String
    
    On Error GoTo errores
    
    m_HTMLTablaProyectoDatos = Me.Proyecto.HTMLTablaProyectoDatos
    p_Error = Me.Proyecto.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaProyectoDatos = Replace(m_HTMLTablaProyectoDatos, "<table>", "<table class='action-table'>")
    
    m_HTMLCuerpo = ""
    m_HTMLCuerpo = m_HTMLCuerpo & "<div class='card'>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<p>Se ha solicitado la publicación de la edición <strong>" & HTMLSafe(Nz(Me.Edicion, "")) & "</strong> del proyecto <strong>" & HTMLSafe(Nz(Me.Proyecto.NombreProyecto, Nz(Me.Proyecto.Proyecto, ""))) & "</strong>.</p>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "</div>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<br /><br />" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<div class='card'>" & m_HTMLTablaProyectoDatos & "</div>" & vbNewLine
    getHTMLEdicionPropuestaPublicacion = CuerpoHTMLConEstiloCorporativo("Propuesta de Publicación", m_HTMLCuerpo)
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mï¿½todo getHTMLEdicionPropuestaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getHTMLEdicionRevision( _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_HTMLUsuariosConectados As String
    Dim m_HTMLTablaProyectoDatos As String
    Dim m_HTMLCuerpo As String
    On Error GoTo errores
    m_HTMLUsuariosConectados = getHTMLUsuariosConectados(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaProyectoDatos = Me.Proyecto.HTMLTablaProyectoDatos
    p_Error = Me.Proyecto.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaProyectoDatos = Replace(m_HTMLTablaProyectoDatos, "<table>", "<table class='action-table'>")
    m_HTMLCuerpo = ""
    m_HTMLCuerpo = m_HTMLCuerpo & "<div class='card'>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & m_HTMLUsuariosConectados & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<br /><br />" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<strong>CALIDAD TE PROPONE QUE REVISES LA EDICIÓN</strong>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "</div>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<br /><br />" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<div class='card'>" & m_HTMLTablaProyectoDatos & "</div>" & vbNewLine
    getHTMLEdicionRevision = CuerpoHTMLConEstiloCorporativo("Propuesta de Revisión", m_HTMLCuerpo)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLEdicionRevision ha devuelto el error: " & Err.Description
    End If
End Function
Private Function getHTMLEdicionQuitarPropuestaPublicacion( _
                                                            Optional ByRef p_Error As String _
                                                            ) As String
    Dim m_HTMLUsuariosConectados As String
    Dim m_HTMLTablaProyectoDatos As String
    Dim m_HTMLCuerpo As String
    On Error GoTo errores
    m_HTMLUsuariosConectados = getHTMLUsuariosConectados(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaProyectoDatos = Me.Proyecto.HTMLTablaProyectoDatos
    p_Error = Me.Proyecto.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaProyectoDatos = Replace(m_HTMLTablaProyectoDatos, "<table>", "<table class='action-table'>")
    m_HTMLCuerpo = ""
    m_HTMLCuerpo = m_HTMLCuerpo & "<div class='card'>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & m_HTMLUsuariosConectados & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<br /><br />" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<strong>PROPUESTA DE PUBLICACIÓN ELIMINADA</strong>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "</div>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<br /><br />" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<div class='card'>" & m_HTMLTablaProyectoDatos & "</div>" & vbNewLine
    getHTMLEdicionQuitarPropuestaPublicacion = CuerpoHTMLConEstiloCorporativo("Quitar Propuesta de Publicación", m_HTMLCuerpo)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLEdicionQuitarPropuestaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getHTMLEdicionRechazarPropuestaPublicacion( _
                                                            Optional ByRef p_Error As String _
                                                            ) As String
    Dim m_HTMLUsuariosConectados As String
    Dim m_HTMLTablaProyectoDatos As String
    Dim m_HTMLCuerpo As String
    On Error GoTo errores
    m_HTMLUsuariosConectados = getHTMLUsuariosConectados(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaProyectoDatos = Me.Proyecto.HTMLTablaProyectoDatos
    p_Error = Me.Proyecto.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaProyectoDatos = Replace(m_HTMLTablaProyectoDatos, "<table>", "<table class='action-table'>")
    m_HTMLCuerpo = ""
    m_HTMLCuerpo = m_HTMLCuerpo & "<div class='card'>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & m_HTMLUsuariosConectados & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<br /><br />" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<strong>PROPUESTA DE PUBLICACIÓN RECHAZADA POR CALIDAD</strong>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<br /><br />" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & HTMLSafeLargo(Me.PropuestaRechazadaPorCalidadMotivo) & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<br /><br />" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "Ha de cambiar lo que está mal y volver a proponer la publicación" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "</div>" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<br /><br />" & vbNewLine
    m_HTMLCuerpo = m_HTMLCuerpo & "<div class='card'>" & m_HTMLTablaProyectoDatos & "</div>" & vbNewLine
    getHTMLEdicionRechazarPropuestaPublicacion = CuerpoHTMLConEstiloCorporativo("Rechazo Propuesta de Publicación", m_HTMLCuerpo)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getHTMLEdicionRechazarPropuestaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function


Public Function getColeEdicionesIgualesEnEdcionesAnteriores( _
                                                            Optional ByRef p_Error As String _
                                                            ) As Collection
    Dim rcdDatos As DAO.Recordset
    Dim m_IDEdicion As String
    Dim m_IDProyecto As String
    Dim m_SQL As String
    
    
    'primero he de conocer el IDProyecto,Riesgo,CodPM
    On Error GoTo errores
    m_IDEdicion = Me.IDEdicion
    m_IDProyecto = Me.IDProyecto
    
    
    If m_IDProyecto = "" Or m_IDEdicion = "" Then
        Exit Function
    End If
    
    m_SQL = "SELECT TbProyectosEdiciones.IDEdicion " & _
            "FROM TbProyectosEdiciones " & _
            "WHERE (((TbProyectosEdiciones.IDProyecto)=" & m_IDProyecto & _
            ") AND ((TbProyectosEdiciones.IDEdicion)<=" & m_IDEdicion & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                m_IDEdicion = .Fields("IDEdicion")
                If getColeEdicionesIgualesEnEdcionesAnteriores Is Nothing Then
                    Set getColeEdicionesIgualesEnEdcionesAnteriores = New Collection
                End If
                getColeEdicionesIgualesEnEdcionesAnteriores.Add m_IDEdicion
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColeEdicionesIgualesEnEdcionesAnteriores ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CerrarRiesgos( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_SQL As String
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Riesgo As riesgo
    Dim m_Estado As String
    Dim m_FechaRef As Date
    
    'primero he de conocer el IDProyecto,Riesgo,CodPM
    On Error GoTo errores
    m_FechaRef = Date
    CerrarAcciones p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    'm_FechaRef = CDate("12/12/2023")
    Set m_Col = Me.colRiesgos
    
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_Riesgo = m_Col(m_ID)
            m_Estado = m_Riesgo.ESTADOCalculadoTexto
            If m_Riesgo.Estado <> "Retirado" Then
                m_SQL = "UPDATE TbRiesgos SET Estado = '" & m_Estado & _
                        "', FechaEstado = '" & m_FechaRef & "' " & "WHERE IDRiesgo=" & _
                        m_Riesgo.IDRiesgo & ";"
                InvalidarCacheRiesgo m_Riesgo.IDRiesgo, p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                getdb().Execute m_SQL
            End If
            Set m_Riesgo = Nothing
        Next
    End If
     
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CerrarRiesgos ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CerrarAccionesCrearReversas( _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    
    Dim m_AccionMitigacionReversa As PMAccionReversa
    Dim m_AccionContingenciaReversa As PCAccionReversa
    
    On Error GoTo errores

    InvalidarCacheEdicion Me.IDEdicion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
    m_SQL = "SELECT TbRiesgosPlanMitigacionDetalle.* " & _
            "FROM ((TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion) " & _
            "INNER JOIN TbRiesgosPlanMitigacionPpal ON TbRiesgos.IDRiesgo = TbRiesgosPlanMitigacionPpal.IDRiesgo) " & _
            "INNER JOIN TbRiesgosPlanMitigacionDetalle " & _
            "ON TbRiesgosPlanMitigacionPpal.IDMitigacion = TbRiesgosPlanMitigacionDetalle.IDMitigacion " & _
            "WHERE (((TbProyectosEdiciones.IDEdicion)=" & Me.IDEdicion & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
                Do While Not .EOF
                    If IsDate(.Fields("FechaInicio")) Or IsDate(.Fields("FechaFinPrevista")) Or IsDate(.Fields("FechaFinReal")) Then
                        Set m_AccionMitigacionReversa = New PMAccionReversa
                        m_AccionMitigacionReversa.IDAccionMitigacion = .Fields("IDAccionMitigacion")
                        m_AccionMitigacionReversa.IDEdicion = Me.IDEdicion
                        m_AccionMitigacionReversa.FechaInicio = Nz(.Fields("FechaInicio"), "")
                        m_AccionMitigacionReversa.FechaFinPrevista = Nz(.Fields("FechaFinPrevista"), "")
                        m_AccionMitigacionReversa.FechaFinReal = Nz(.Fields("FechaFinReal"), "")
                        m_AccionMitigacionReversa.Registrar p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        Set m_AccionMitigacionReversa = Nothing
                    End If
                    .MoveNext
                Loop
            
        End If
    End With
    
   m_SQL = "SELECT TbRiesgosPlanContingenciaDetalle.* " & _
            "FROM ((TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion) " & _
            "INNER JOIN TbRiesgosPlanContingenciaPpal ON TbRiesgos.IDRiesgo = TbRiesgosPlanContingenciaPpal.IDRiesgo) " & _
            "INNER JOIN TbRiesgosPlanContingenciaDetalle " & _
            "ON TbRiesgosPlanContingenciaPpal.IDContingencia = TbRiesgosPlanContingenciaDetalle.IDContingencia " & _
            "WHERE (((TbProyectosEdiciones.IDEdicion)=" & Me.IDEdicion & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                If IsDate(.Fields("FechaInicio")) Or IsDate(.Fields("FechaFinPrevista")) Or IsDate(.Fields("FechaFinReal")) Then
                    Set m_AccionContingenciaReversa = New PCAccionReversa
                    m_AccionContingenciaReversa.IDAccionContingencia = .Fields("IDAccionContingencia")
                    m_AccionContingenciaReversa.IDEdicion = Me.IDEdicion
                    m_AccionContingenciaReversa.FechaInicio = Nz(.Fields("FechaInicio"), "")
                    m_AccionContingenciaReversa.FechaFinPrevista = Nz(.Fields("FechaFinPrevista"), "")
                    m_AccionContingenciaReversa.FechaFinReal = Nz(.Fields("FechaFinReal"), "")
                    m_AccionContingenciaReversa.Registrar p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Set m_AccionContingenciaReversa = Nothing
                End If
                .MoveNext
            Loop
            
        End If
    End With
    
     
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CerrarAcciones ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CerrarAccionesDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    
    Dim m_AccionMitigacionReversa As PMAccionReversa
    Dim m_AccionContingenciaReversa As PCAccionReversa
    
    On Error GoTo errores

    InvalidarCacheEdicion Me.IDEdicion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
    m_SQL = "SELECT TbRiesgosPlanMitigacionDetalle.* " & _
            "FROM ((TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion) " & _
            "INNER JOIN TbRiesgosPlanMitigacionPpal ON TbRiesgos.IDRiesgo = TbRiesgosPlanMitigacionPpal.IDRiesgo) " & _
            "INNER JOIN TbRiesgosPlanMitigacionDetalle " & _
            "ON TbRiesgosPlanMitigacionPpal.IDMitigacion = TbRiesgosPlanMitigacionDetalle.IDMitigacion " & _
            "WHERE (((TbProyectosEdiciones.IDEdicion)=" & Me.IDEdicion & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                If IsDate(.Fields("FechaFinPrevista")) And Not IsDate(.Fields("FechaFinReal")) Then
                    .Edit
                        .Fields("FechaFinReal") = Date
                    .Update
                End If
                .MoveNext
            Loop
                
        End If
    End With
    
   m_SQL = "SELECT TbRiesgosPlanContingenciaDetalle.* " & _
            "FROM ((TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion) " & _
            "INNER JOIN TbRiesgosPlanContingenciaPpal ON TbRiesgos.IDRiesgo = TbRiesgosPlanContingenciaPpal.IDRiesgo) " & _
            "INNER JOIN TbRiesgosPlanContingenciaDetalle " & _
            "ON TbRiesgosPlanContingenciaPpal.IDContingencia = TbRiesgosPlanContingenciaDetalle.IDContingencia " & _
            "WHERE (((TbProyectosEdiciones.IDEdicion)=" & Me.IDEdicion & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
                Do While Not .EOF
                If IsDate(.Fields("FechaFinPrevista")) And Not IsDate(.Fields("FechaFinReal")) Then
                    .Edit
                        .Fields("FechaFinReal") = Date
                    .Update
                End If
                .MoveNext
            Loop
        End If
    End With
    
     
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CerrarAccionesDatos ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CerrarAcciones( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
   
    
    On Error GoTo errores
   
    CerrarAccionesCrearReversas p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    CerrarAccionesDatos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CerrarAcciones ha devuelto el error: " & Err.Description
    End If
End Function
Public Function AbrirAcciones( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
   
    
    On Error GoTo errores
   
    AbrirAccionesDatos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    AbrirAccionesEliminarReversas p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbrirAcciones ha devuelto el error: " & Err.Description
    End If
End Function
Public Function AbrirAccionesDatos( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    
    Dim m_AccionMitigacionReversa As PMAccionReversa
    Dim m_AccionContingenciaReversa As PCAccionReversa
    
    On Error GoTo errores

    InvalidarCacheEdicion Me.IDEdicion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
    m_SQL = "SELECT TbRiesgosPlanMitigacionDetalle.* " & _
            "FROM ((TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion) " & _
            "INNER JOIN TbRiesgosPlanMitigacionPpal ON TbRiesgos.IDRiesgo = TbRiesgosPlanMitigacionPpal.IDRiesgo) " & _
            "INNER JOIN TbRiesgosPlanMitigacionDetalle " & _
            "ON TbRiesgosPlanMitigacionPpal.IDMitigacion = TbRiesgosPlanMitigacionDetalle.IDMitigacion " & _
            "WHERE (((TbProyectosEdiciones.IDEdicion)=" & Me.IDEdicion & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                If IsDate(.Fields("FechaFinReal")) Then
                    Set m_AccionMitigacionReversa = Constructor.getPmAccionReversa(.Fields("IDAccionMitigacion"), p_Error)
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If Not m_AccionMitigacionReversa Is Nothing Then
                        If IsDate(m_AccionMitigacionReversa.FechaFinReal) Then
                            .Edit
                                .Fields("FechaFinReal") = m_AccionMitigacionReversa.FechaFinReal
                            .Update
                        Else
                            .Edit
                                .Fields("FechaFinReal") = Null
                            .Update
                        End If
                    End If
                    
                End If
                Set m_AccionMitigacionReversa = Nothing
                .MoveNext
            Loop
                
        End If
    End With
    
   m_SQL = "SELECT TbRiesgosPlanContingenciaDetalle.* " & _
            "FROM ((TbProyectosEdiciones INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion) " & _
            "INNER JOIN TbRiesgosPlanContingenciaPpal ON TbRiesgos.IDRiesgo = TbRiesgosPlanContingenciaPpal.IDRiesgo) " & _
            "INNER JOIN TbRiesgosPlanContingenciaDetalle " & _
            "ON TbRiesgosPlanContingenciaPpal.IDContingencia = TbRiesgosPlanContingenciaDetalle.IDContingencia " & _
            "WHERE (((TbProyectosEdiciones.IDEdicion)=" & Me.IDEdicion & "));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                If IsDate(.Fields("FechaFinReal")) Then
                    Set m_AccionContingenciaReversa = Constructor.getPcAccionReversa(.Fields("IDAccionContingencia"), p_Error)
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If Not m_AccionContingenciaReversa Is Nothing Then
                        If IsDate(m_AccionContingenciaReversa.FechaFinReal) Then
                            .Edit
                                .Fields("FechaFinReal") = m_AccionContingenciaReversa.FechaFinReal
                            .Update
                        Else
                            .Edit
                                .Fields("FechaFinReal") = Null
                            .Update
                        End If
                    End If
                    
                End If
                Set m_AccionContingenciaReversa = Nothing
                
                .MoveNext
            Loop
                
        End If
    End With
    
     
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbrirAccionesDatos ha devuelto el error: " & Err.Description
    End If
End Function
Public Function AbrirAccionesEliminarReversas( _
                                                Optional ByRef p_Error As String _
                                                ) As String
                    
    Dim m_SQL As String
    
    On Error GoTo errores
    
    m_SQL = "DELETE * " & _
            "FROM TbRiesgosPlanMitigacionDetalleReversa " & _
            "WHERE IDEdicion=" & Me.IDEdicion & ";"
    
    getdb().Execute m_SQL
    
    m_SQL = "DELETE * " & _
            "FROM TbRiesgosPlanContingenciaDetalleReversa " & _
            "WHERE IDEdicion=" & Me.IDEdicion & ";"
    
    getdb().Execute m_SQL
     
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbrirAccionesEliminarReversas ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
    Dim m_ID As String
    Dim m_ObjProyecto As Proyecto
    Dim m_ObjDocumento As Documento
    Dim m_ObjEdicionUltima As Edicion
    Dim m_ObjEdicionUltimaAlInicio As Edicion
    Dim m_IDEdicion As String
    Dim m_SQL As String
    On Error GoTo errores
    
    p_Error = ""
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_IDEdicion = Me.IDEdicion
    Set m_ObjDocumento = Me.DocumentoObj
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjDocumento Is Nothing Then
        If fso.FileExists(m_ObjDocumento.URLAdjunto) Then
            If FicheroAbierto(m_ObjDocumento.URLAdjunto) Then
                p_Error = "Tiene abierto el informe de la edición"
                Err.Raise 1000
            End If
        End If
        m_ObjDocumento.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    m_SQL = "DELETE * FROM TbProyectosEdiciones " & _
            "WHERE IDEdicion=" & Me.IDEdicion & ";"
    getdb().Execute m_SQL
    Set m_ObjProyecto = Constructor.getProyecto(Me.IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set m_ObjEdicionUltima = m_ObjProyecto.EdicionUltima
    Set m_ObjEdicionUltimaAlInicio = Constructor.getEdicion(m_ObjEdicionUltima.IDEdicion)
    Set m_ObjDocumento = m_ObjEdicionUltima.DocumentoObj
    p_Error = m_ObjEdicionUltima.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjDocumento Is Nothing Then
        m_ObjDocumento.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If Not m_ObjEdicionUltima Is Nothing Then
        With m_ObjEdicionUltima
            .FechaPreparadaParaPublicar = ""
            .FechaPublicacion = ""
            .IDDocumentoAGEDO = ""
            .Registrar m_ObjEdicionUltimaAlInicio, EnumSiNo.Sí, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End With
    End If
    Set m_ObjEdicionUltima.Proyecto = Nothing
    EliminarPublicacionesLog m_IDEdicion
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Borrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function EstablecerFechaMaxPublicacion( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim m_SQL As String
    Dim m_FechaMaxProximaPublicacion As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    p_Error = ""
    m_FechaMaxProximaPublicacion = Me.Proyecto.FechaMaxProximaPublicacionCalculada
    
    
    
    
    If IsDate(m_FechaMaxProximaPublicacion) Then
        m_SQL = "UPDATE TbProyectosEdiciones SET FechaMaxProximaPublicacion = #" & m_FechaMaxProximaPublicacion & "# " & _
                "WHERE (((TbProyectosEdiciones.IDEdicion)=" & Me.IDEdicion & "));"
    Else
        m_SQL = "UPDATE TbProyectos SET TbProyectos.FechaMaxProximaPublicacion = Null WHERE (((TbProyectos.IDProyecto)=" & Me.IDProyecto & "));"
    End If
    getdb().Execute m_SQL
    Exit Function
    
    
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerFechaMaxPublicacion ha devuelto el error: " & Err.Description
    End If
    
End Function
Private Function MotivoNoOK( _
                                Optional ByRef p_ObjEdicionAlInicio As Edicion, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim m_valor As String
    Dim m_UsuarioAprueba As String
    Dim m_UsuarioRevisa As String
    Dim m_UsuarioElabora As String
    Dim m_NombreCampo As String
    
    On Error GoTo errores
    
    
    If Me.Proyecto Is Nothing Then
        MotivoNoOK = "No se conoce la gestión de riesgos"
        Exit Function
    End If

    If Me.UsuarioConectadoAutorizado = EnumSiNo.No Then
        MotivoNoOK = "El perfil de su usuario no tiene autorización para realizar esta operación"
        Exit Function
    End If
    If Not IsDate(Me.FechaEdicion) Then
        MotivoNoOK = "Falta Fecha Edición"
        Exit Function
    End If
    If Me.Edicion = "" Then
        MotivoNoOK = "Falta Edición"
        Exit Function
    End If
    If Not Me.Proyecto.EdicionUltima Is Nothing Then
        If CInt(Me.Proyecto.EdicionUltima.Edicion) > CInt(Me.Edicion) Then
            MotivoNoOK = "El número de la edición ha de ser mayor que la última"
            Exit Function
        End If
    Else
        If CInt(Me.Edicion) <> 1 Then
            MotivoNoOK = "El número de la edición ha de ser 1 ya que es la primera"
            Exit Function
        End If
    End If
    If Me.Elaborado = "" Then
        MotivoNoOK = "Falta Elaborado"
        Exit Function
    End If
    If Me.Revisado = "" Then
        MotivoNoOK = "Falta Revisado"
        Exit Function
    End If
    If Me.Aprobado = "" Then
        MotivoNoOK = "Falta Aprobado"
        Exit Function
    End If
    If Me.Elaborado = Me.Revisado Then
        MotivoNoOK = "El Usuario que Elabora no puede ser el mismo que revisa"
        Exit Function
    End If
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
     End If
    
End Function
Public Function Registrar( _
                            Optional ByRef p_ObjEdicionAlInicio As Edicion, _
                            Optional p_ProvieneDePublicacion As EnumSiNo = EnumSiNo.No, _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    Dim rcdDatos As DAO.Recordset
     Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim m_fechaMaxPublicacion As String
    Dim m_FechaRef As Date
    On Error GoTo errores
    
    Me.Error = ""
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    If p_ProvieneDePublicacion = 0 Then
        p_ProvieneDePublicacion = EnumSiNo.No
    End If
    
    m_MotivoNoOK = MotivoNoOK(p_ObjEdicionAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_ObjEdicionAlInicio Is Nothing Then
        Me.IDEdicion = Me.IDEdicionCalculada
        If Me.Proyecto.EdicionUltima Is Nothing Then
            Me.Edicion = "1"
        Else
            Me.Edicion = CStr(CLng(Me.Proyecto.EdicionUltima.Edicion) + 1)
        End If
        m_SQL = "TbProyectosEdiciones"
        
        
    Else
        If Me.IDEdicion = "" Then
            p_Error = "Falta el IDEdición"
            Err.Raise 1000
        End If
        m_SQL = "SELECT * FROM TbProyectosEdiciones " & _
                "WHERE IDEdicion=" & Me.IDEdicion & ";"
        
    End If
    If p_ProvieneDePublicacion = EnumSiNo.No Then
        If Me.FechaMaxProximaPublicacion = "" Then
            m_fechaMaxPublicacion = Me.FechaMaxProximaPublicacionCalculado
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    Else
        m_fechaMaxPublicacion = DateAdd("m", CDbl(m_ObjEntorno.JPMesesAvisoEntreEdiciones), m_FechaRef)
    End If
    Me.FechaMaxProximaPublicacion = m_fechaMaxPublicacion
    InvalidarCacheEdicion Me.IDEdicion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If p_ObjEdicionAlInicio Is Nothing Then
            .AddNew
                .Fields("IDEdicion") = Me.IDEdicion
                .Fields("Edicion") = Me.Edicion
                .Fields("IDProyecto") = Me.IDProyecto
                .Fields("FechaEdicion") = Me.FechaEdicion
                .Fields("Elaborado") = Me.Elaborado
                .Fields("Revisado") = Me.Revisado
                .Fields("Aprobado") = Me.Aprobado
                If IsDate(Me.FechaPublicacion) Then
                    .Fields("FechaPublicacion") = Me.FechaPublicacion
                End If
                If Me.Comentarios <> "" Then
                    .Fields("Comentarios") = Me.Comentarios
                End If
                If Me.PermitidoImprimirExcel <> "" Then
                    .Fields("PermitidoImprimirExcel") = Me.PermitidoImprimirExcel
                End If
                If Me.IDDocumentoAGEDO <> "" Then
                    .Fields("IDDocumentoAGEDO") = Me.IDDocumentoAGEDO
                End If
                If Me.NombreArchivoInforme <> "" Then
                    .Fields("NombreArchivoInforme") = Me.NombreArchivoInforme
                End If
                If Me.FechaMaxProximaPublicacion <> "" Then
                    .Fields("FechaMaxProximaPublicacion") = Me.FechaMaxProximaPublicacion
                End If
                If Me.UsuarioCalidadRechazaPropuesta <> "" Then
                    .Fields("UsuarioCalidadRechazaPropuesta") = Me.UsuarioCalidadRechazaPropuesta
                End If
            .Update
        Else
            .Edit
                
                .Fields("FechaEdicion") = Me.FechaEdicion
                .Fields("Elaborado") = Me.Elaborado
                .Fields("Revisado") = Me.Revisado
                .Fields("Aprobado") = Me.Aprobado
                If IsDate(Me.FechaPublicacion) Then
                    .Fields("FechaPublicacion") = Me.FechaPublicacion
                Else
                    .Fields("FechaPublicacion") = Null
                End If
                If Me.Comentarios <> "" Then
                    .Fields("Comentarios") = Me.Comentarios
                Else
                    .Fields("Comentarios") = Null
                End If
                If Me.PermitidoImprimirExcel <> "" Then
                    .Fields("PermitidoImprimirExcel") = Me.PermitidoImprimirExcel
                Else
                    .Fields("PermitidoImprimirExcel") = Null
                End If
                If Me.IDDocumentoAGEDO <> "" Then
                    .Fields("IDDocumentoAGEDO") = Me.IDDocumentoAGEDO
                Else
                    .Fields("IDDocumentoAGEDO") = Null
                End If
                If Me.NombreArchivoInforme <> "" Then
                    .Fields("NombreArchivoInforme") = Me.NombreArchivoInforme
                Else
                    .Fields("NombreArchivoInforme") = Null
                End If
                If Me.FechaMaxProximaPublicacion <> "" Then
                    .Fields("FechaMaxProximaPublicacion") = Me.FechaMaxProximaPublicacion
                Else
                    .Fields("FechaMaxProximaPublicacion") = Null
                End If
                If Me.UsuarioCalidadRechazaPropuesta <> "" Then
                    .Fields("UsuarioCalidadRechazaPropuesta") = Me.UsuarioCalidadRechazaPropuesta
                Else
                    .Fields("UsuarioCalidadRechazaPropuesta") = Null
                End If
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CambiarFechaProximaPublicacion( _
                                                p_EdicionAlinicio As Edicion, _
                                                Optional ByRef p_Error As String _
                                                ) As String
                            
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    If p_EdicionAlinicio Is Nothing Then
        p_Error = "No se sabe la edición al inicio"
        Err.Raise 1000
    End If
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "No se puede cambiar la fecha de proxima publicación a una edición no activa"
        Err.Raise 1000
    End If
    If Not IsDate(Nz(Me.FechaMaxProximaPublicacion, "")) Then
        p_Error = "Ha de introducir una fecha válida"
        Err.Raise 1000
    End If
    With p_EdicionAlinicio
        If Format(.FechaMaxProximaPublicacion, "dd/mm/yyyy") = Format(Me.FechaMaxProximaPublicacion, "dd/mm/yyyy") Then
            m_Error = "No se ha cambiado la fecha"
            Err.Raise 1000
        End If
    End With
    
    m_SQL = "UPDATE TbProyectosEdiciones SET FechaMaxProximaPublicacion = #" & _
            Format(Me.FechaMaxProximaPublicacion, "mm/dd/yyyy") & "# " & _
            "WHERE IDEdicion=" & Me.IDEdicion & ";"
    getdb().Execute m_SQL
    
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CambiarFechaProximaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function QuitarPublicacion( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                
    Dim m_ObjProyectoAlInicio As Proyecto
   
    Dim m_IDDocumento As String
    Dim m_NombreDocumento As String
    Dim m_ObjDocumento As Documento
    
    On Error GoTo errores
    
    With Me
        .FechaPublicacion = ""
        m_IDDocumento = .IDDocumentoAGEDO
        m_NombreDocumento = .NombreArchivoInforme
        If m_IDDocumento <> "" Then
            Set m_ObjDocumento = Constructor.getDocumentoByID(m_IDDocumento, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            m_ObjDocumento.Borrar p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If m_NombreDocumento <> "" Then
            If fso.FileExists(.URLArchivoInforme) Then
                fso.DeleteFile .URLArchivoInforme, True
            End If
        End If
        .NombreArchivoInforme = ""
        .IDDocumentoAGEDO = ""
        .Registrar , EnumSiNo.No, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método QuitarPublicacion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function ProponerParaPublicacion( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ColLineas As Scripting.Dictionary
    Dim m_PublicacionLog As PublicacionLog
    
    
    On Error GoTo errores
    
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "La Edición no es susceptible de ser publicada"
        Err.Raise 1000
    End If
    
    If Me.Publicable = EnumSiNo.No Then
        p_Error = "No ha superado el informe de publicabilidad"
        Err.Raise 1000
    End If
    
    
    m_SQL = "SELECT  FechaPreparadaParaPublicar, PropuestaRechazadaPorCalidadFecha, " & _
            "PropuestaRechazadaPorCalidadMotivo,UsuarioProponePublicar " & _
            "FROM TbProyectosEdiciones " & _
            "WHERE IDEdicion=" & Me.IDEdicion & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            Me.FechaPreparadaParaPublicar = Date
            Me.PropuestaRechazadaPorCalidadFecha = ""
            Me.PropuestaRechazadaPorCalidadMotivo = ""
            Me.UsuarioProponePublicar = m_ObjUsuarioConectado.UsuarioRed
            
            .Edit
                .Fields("FechaPreparadaParaPublicar") = Me.FechaPreparadaParaPublicar
                .Fields("PropuestaRechazadaPorCalidadFecha") = Null
                .Fields("PropuestaRechazadaPorCalidadMotivo") = Null
                .Fields("UsuarioProponePublicar") = Me.UsuarioProponePublicar
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Set m_PublicacionLog = New PublicacionLog
    With m_PublicacionLog
        .IDEdicion = Me.IDEdicion
        
        
        .FechaPreparadaParaPublicar = Me.FechaPreparadaParaPublicar
        .UsuarioFechaPreparadaParaPublicar = m_ObjUsuarioConectado.Nombre
        .Registrar
    End With
    Set m_PublicacionLog = Nothing
    EnviarCorreoPropuestaNuevaPublicacion Me, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        EstablecerTareasTecnico EnumSiNo.Sí, m_Error
    Else
        EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
    End If
    
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ProponerParaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function
Public Function QuitarPropuestaParaPublicacion( _
                                                Optional ByRef p_Error As String _
                                                ) As String
                    
    Dim m_Proyecto As Proyecto
    Dim m_SQL As String
    Dim m_PublicacionLog As PublicacionLog
    
    On Error GoTo errores
    Set m_Proyecto = Me.Proyecto
    
    
    If Me.Proyecto.EsActivo = EnumSiNo.No Then
        p_Error = Me.Proyecto.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "El proyecto no es activo"
        Err.Raise 1000
    End If
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "La Edición no es susceptible de ser publicada"
        Err.Raise 1000
    End If
    
        
    Me.FechaPreparadaParaPublicar = ""
    m_SQL = "UPDATE TbProyectosEdiciones SET FechaPreparadaParaPublicar = Null " & _
            ", UsuarioProponePublicar = NULL " & _
            "WHERE IDEdicion=" & Me.IDEdicion & ";"
    getdb().Execute m_SQL
    EnviarCorreoQuitarPropuestaNuevaPublicacion Me, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_PublicacionLog = New PublicacionLog
    With m_PublicacionLog
        .IDEdicion = Me.IDEdicion
        .FechaPreparadaParaPublicarQuitar = Now()
        .UsuarioFechaPreparadaParaPublicarQuitar = m_ObjUsuarioConectado.Nombre
        .Registrar
    End With
    Set m_PublicacionLog = Nothing
    If EsTecnico = EnumSiNo.Sí Then
        EstablecerTareasTecnico EnumSiNo.Sí, p_Error
    Else
        EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, p_Error
    End If
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método QuitarPropuestaParaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function RechazoPropuestaParaPublicacion( _
                                                Optional ByRef p_Error As String _
                                                ) As CORREO
                    
    
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    Dim m_PublicacionLog As PublicacionLog
    Dim m_CorreoRechazo As CORREO
    Dim m_EdicionCorreoRevision As EdicionCorreoRevision
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"

    ' Fix #68 (issue #64 cluster): defense-in-depth guard. The menu gate already
    ' restricts this flow to calidad users, but a VBA bypass (Immediate Window,
    ' future automated script) would skip the menu check. Verify the global role
    ' flag BEFORE any DB write so a non-calidad caller cannot persist a rechazo
    ' nor queue a correo. The check is intentionally against the global EsCalidad
    ' (set during session start in Variables Globales.bas), not against
    ' m_ObjUsuarioConectado.EsCalidadCalculado, so the guard works even when the
    ' connected user object is Nothing (test scenarios).
    If EsCalidad <> EnumSiNo.Sí Then
        p_Error = "El usuario no tiene permisos para rechazar propuestas (no es Calidad)."
        Err.Raise 1000
    End If

    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "La Edición no es susceptible de ser publicada"
        Err.Raise 1000
    End If
    If Not IsDate(Me.FechaPreparadaParaPublicar) Then
        p_Error = "La Edición no aparece como propuesta par publicación"
        Err.Raise 1000
    End If
    If Me.PropuestaRechazadaPorCalidadMotivo = "" Then
        p_Error = "No se encuentran los motivos por los que ha sido rechazada la propuesta de publicación"
        Err.Raise 1000
    End If
    Me.PropuestaRechazadaPorCalidadFecha = CStr(m_FechaRef)
    Me.UsuarioCalidadRechazaPropuesta = m_ObjUsuarioConectado.UsuarioRed
    m_SQL = "SELECT * FROM TbProyectosEdiciones " & _
            "WHERE IDEdicion=" & Me.IDEdicion & ";"
       
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "La Edición no aparece estar registrada"
            Err.Raise 1000
        End If
        
        .Edit
            .Fields("PropuestaRechazadaPorCalidadFecha") = Me.PropuestaRechazadaPorCalidadFecha
            .Fields("PropuestaRechazadaPorCalidadMotivo") = Me.PropuestaRechazadaPorCalidadMotivo
            .Fields("PropuestaRechazadaPorCalidadMotivo") = Me.PropuestaRechazadaPorCalidadMotivo
            Me.FechaPreparadaParaPublicar = ""
            .Fields("FechaPreparadaParaPublicar") = Null
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Set m_CorreoRechazo = EnviarCorreoRechazoPropuestaPublicacion(Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_CorreoRechazo Is Nothing Then
        If m_CorreoRechazo.IDCorreo <> "" Then
            Set m_EdicionCorreoRevision = New EdicionCorreoRevision
            m_EdicionCorreoRevision.RegistrarCorreoEnviado Me.IDEdicion, m_CorreoRechazo.IDCorreo, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_EdicionCorreoRevision = Nothing
        End If
    End If
    Set m_PublicacionLog = New PublicacionLog
    With m_PublicacionLog
        .IDEdicion = Me.IDEdicion
        .PropuestaRechazadaPorCalidadFecha = Me.PropuestaRechazadaPorCalidadFecha
        .PropuestaRechazadaPorCalidadMotivo = Me.PropuestaRechazadaPorCalidadMotivo
        .Registrar
    End With
    Set m_PublicacionLog = Nothing
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    Set RechazoPropuestaParaPublicacion = m_CorreoRechazo
    Set m_CorreoRechazo = Nothing

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RechazoPropuestaParaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function
Public Function RegistrarNotasParaCalidad( _
                                                Optional ByRef p_Error As String _
                                                ) As String
                    
    
    Dim m_SQL As String
    
    
    
    On Error GoTo errores
   
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = "La Edición no es activa"
        Err.Raise 1000
    End If
    If Me.NotasCalidadParaPublicar = "" Then
        m_SQL = "UPDATE TbProyectosEdiciones SET NotasCalidadParaPublicar = Null " & _
                "WHERE IDEdicion=" & Me.IDEdicion & ";"
    Else
        m_SQL = "UPDATE TbProyectosEdiciones SET NotasCalidadParaPublicar = '" & Me.NotasCalidadParaPublicar & "' " & _
                "WHERE IDEdicion=" & Me.IDEdicion & ";"
    End If
    getdb().Execute m_SQL
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistrarNotasParaCalidad ha devuelto el error: " & Err.Description
    End If
End Function


Public Function PublicableResetear( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    
    On Error GoTo errores
    m_sURLArchivoInforme = ""
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PublicableResetear ha devuelto el error: " & Err.Description
    End If
End Function

Public Function PriorizacionQuitar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_SQL As String
    Dim m_ColRiesgos As Scripting.Dictionary
    Dim m_ID As Variant
    
    On Error GoTo errores
    
    If Me.EsActivo <> EnumSiNo.Sí Then
        p_Error = "La edición no es activa"
        Err.Raise 1000
    End If
    If Me.UsuarioConectadoAutorizado <> EnumSiNo.Sí Then
        p_Error = "Usuario no autorizado a esa operación"
        Err.Raise 1000
    End If
    Set m_ColRiesgos = Me.colRiesgosNoRetirados
    If m_ColRiesgos Is Nothing Then
        Exit Function
    End If
    InvalidarCacheEdicion Me.IDEdicion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    For Each m_ID In m_ColRiesgos
        m_SQL = "UPDATE TbRiesgos SET Priorizacion = Null " & _
            "WHERE IDRiesgo=" & m_ID & ";"
        
        getdb().Execute m_SQL
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PriorizacionQuitar ha devuelto el error: " & Err.Description
    End If
End Function


Public Function Cerrar( _
                        Optional p_FechaPublicacion As String, _
                        Optional ByRef p_Error As String _
                        ) As String
    'función que va a recorrer todos los riesgos no retirados ni aceptados y los va a poner como cerrados
    
    Dim m_IdRiesgo As Variant
    Dim m_SQL As String
    Dim m_EstadoActual As EnumRiesgoEstado
    Dim m_ObjRiesgo As riesgo
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    If Not IsDate(p_FechaPublicacion) Then
        p_FechaPublicacion = CStr(m_FechaRef)
    End If
    InvalidarCacheEdicion Me.IDEdicion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For Each m_IdRiesgo In Me.colRiesgos
        If Nz(m_IdRiesgo, "") = "" Then
            GoTo siguienteRiesgo
        End If
        Set m_ObjRiesgo = Me.colRiesgos(m_IdRiesgo)
        With m_ObjRiesgo
            
            m_EstadoActual = .ESTADOCalculado
            If m_EstadoActual <> EnumRiesgoEstado.Retirado And _
                m_EstadoActual <> EnumRiesgoEstado.Aceptado Then
                
                .Estado = "Cerrado"
                .FechaEstado = p_FechaPublicacion
                m_SQL = "UPDATE TbRiesgos SET Estado = '" & .Estado & "', FechaEstado = '" & .FechaEstado & "' " & _
                    "WHERE IDRiesgo=" & m_ObjRiesgo.IDRiesgo & ";"
                InvalidarCacheRiesgo m_ObjRiesgo.IDRiesgo, p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                getdb().Execute m_SQL
            End If
        End With
        
        
        
        Set m_ObjRiesgo = Nothing
siguienteRiesgo:
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Cerrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Abrir( _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim m_IdRiesgo As Variant
    Dim m_Estado As EnumRiesgoEstado
    Dim m_ObjRiesgo As riesgo
    Dim m_FechaPublicacionEdicionAnterior As String
    Dim m_objEdicion As Edicion
    On Error GoTo errores
    
    
    If Me.TieneRiesgos = EnumSiNo.No Then
        Exit Function
    End If
    For Each m_IdRiesgo In Me.colRiesgos.keys
        If Nz(m_IdRiesgo, "") = "" Then
            GoTo siguienteRiesgo
        End If
        Set m_ObjRiesgo = Me.colRiesgos(m_IdRiesgo)
        With m_ObjRiesgo
            m_Estado = .ESTADOCalculado
            
            If m_Estado = EnumRiesgoEstado.Cerrado Then
                Set m_objEdicion = m_ObjRiesgo.Edicion
                p_Error = m_ObjRiesgo.Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If m_objEdicion.EdicionAnterior Is Nothing Then
                    p_Error = m_objEdicion.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    m_FechaPublicacionEdicionAnterior = m_objEdicion.FechaEdicion
                Else
                    p_Error = m_objEdicion.Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    m_FechaPublicacionEdicionAnterior = m_objEdicion.EdicionAnterior.FechaPublicacion
                End If
                
                .FechaEstado = m_FechaPublicacionEdicionAnterior
                .Registrar EnumSiNo.No, , p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                Set m_objEdicion = Nothing
            End If
        End With
        Set m_ObjRiesgo = Nothing
siguienteRiesgo:
    Next
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Abrir ha devuelto el error: " & Err.Description
    End If
End Function
'Public Function ProponerParaPublicacion( _
'                                        Optional ByRef p_Error As String _
'                                        ) As String
'
'    Dim m_Correo As Correo
'    Dim m_SQL As String
'    Dim m_URL As String
'
'    On Error GoTo errores
'    If ColLineasInformePublicabilidad Is Nothing Then
'        'm_URL=
'    End If
'    If Me.Proyecto Is Nothing Then
'        p_Error = Me.Error
'        If p_Error <> "" Then
'            Err.Raise 1000
'        End If
'        p_Error = "No se sabe el Proyecto que es"
'        Err.Raise 1000
'    Else
'        p_Error = Me.Error
'        If p_Error <> "" Then
'            Err.Raise 1000
'        End If
'    End If
'    If Me.Edicion.EsActivo = EnumSiNo.No Then
'        p_Error = Me.Edicion.Error
'        If p_Error <> "" Then
'            Err.Raise 1000
'        End If
'        p_Error = "El proyecto no es activo"
'        Err.Raise 1000
'    End If
'    If Me.EsActivo = EnumSiNo.No Then
'        p_Error = "La Edición no es susceptible de ser publicada"
'        Err.Raise 1000
'    End If
'    Me.FechaPreparadaParaPublicar = Now()
'    m_SQL = "UPDATE TbProyectosEdiciones SET FechaPreparadaParaPublicar = #" & _
'            Format(Me.FechaPreparadaParaPublicar, "mm/dd/yyyy") & _
'            "#, UsuarioProponePublicar = '" & m_ObjUsuarioConectado.UsuarioRed & "' " & _
'            "WHERE IDEdicion=" & Me.IDEdicion & ";"
'    getdb().Execute m_SQL
'    EnviarCorreoPropuestaNuevaPublicacion Me, p_ColLineas, p_Error
'    If p_Error <> "" Then
'        Err.Raise 1000
'    End If
'    EstablecerTareasCalidad EnumTipoTareas.ParaPreparadasParaPublicar, p_Error
'    If p_Error <> "" Then
'        Err.Raise 1000
'    End If
'
'    Exit Function
'
'errores:
'    If Err.Number <> 1000 Then
'        p_Error = "El método ProponerParaPublicacion ha devuelto el error: " & Err.Description
'    End If
'End Function
Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    On Error GoTo errores
    'ha de ser la última Edición
    'ha de tener permiso el que lo ejecuta
    
    
    If Me.UsuarioConectadoAutorizado = EnumSiNo.No Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
        Err.Raise 1000
    End If
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "Sólo se puede borrar la última Edición sin publicar"
        Err.Raise 1000
    End If
    If fso.FileExists(Me.URLArchivoInforme) Then
        If FicheroAbierto(Me.URLArchivoInforme) Then
            p_Error = "Se ha de cerrar el archivo del informe"
            Err.Raise 1000
        End If
    End If
    If Not Me.DocumentoObj Is Nothing Then
        If fso.FileExists(Me.DocumentoObj.URLAdjunto) Then
            If FicheroAbierto(Me.DocumentoObj.URLAdjunto) Then
                p_Error = "Se ha de cerrar el archivo del informe"
                Err.Raise 1000
            End If
        End If
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
                                    
End Function

Public Function EliminarPublicacion( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    
    Dim m_URLInforme As String
    Dim m_SQL As String
    Dim m_ObjDocumento As Documento
    Dim m_NombreArchivoInforme As String
    
    
    On Error GoTo errores
    
    
    Avance "Obteniendo la validación pertinente"
    ValidacionEliminarPublicacion m_URLInforme, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjDocumento = Me.DocumentoObj
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_NombreArchivoInforme = fso.GetFileName(Me.URLArchivoInforme)
    If m_NombreArchivoInforme <> "" Then
        If fso.FileExists(Me.URLArchivoInforme) Then
            If FicheroAbierto(Me.URLArchivoInforme) Then
                p_Error = "Antes de borrar la publicación ha de cerrar el informe"
                Err.Raise 1000
            End If
        End If
    Else
        If Not m_ObjDocumento Is Nothing Then
            If fso.FileExists(m_ObjDocumento.URLAdjunto) Then
                If FicheroAbierto(m_ObjDocumento.URLAdjunto) Then
                    p_Error = "Antes de borrar la publicación ha de cerrar el informe"
                    Err.Raise 1000
                End If
            End If
        End If
    End If
    With Me
        
        .FechaPublicacion = ""
        .IDDocumentoAGEDO = ""
        .NombreArchivoInforme = ""
        m_SQL = "UPDATE TbProyectosEdiciones SET FechaPublicacion = Null, IDDocumentoAGEDO = Null, NombreArchivoInforme = Null " & _
                "WHERE (((TbProyectosEdiciones.IDEdicion)=" & .IDEdicion & "));"
        getdb().Execute m_SQL
        
    End With
    Avance "Datos borrados correctamente"
   
    If Not m_ObjDocumento Is Nothing Then
        m_ObjDocumento.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_NombreArchivoInforme <> "" Then
            If fso.FileExists(Me.URLArchivoInforme) Then
                fso.DeleteFile m_URLInforme, True
            End If
        End If
    End If
    
    RaiseEvent PublicacionEliminada
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método eliminarPublicacion ha devuelto el error: " & Err.Description
    End If
                                    
End Function

Public Function ValidacionEliminarPublicacion( _
                                                Optional ByRef p_URLInforme As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
                                                
    '       1) Ha de estar publicada-->EstaPublicado
    '       2) No ha de haber otra Edición posterior-->EdicionConVersionesPosteriores
    '       3) Ha de estar autorizado
    
    
    On Error GoTo errores
    
    
    With Me
        Avance "Obteniendo la validación pertinente: Usuario Autorizado"
        If .UsuarioConectadoAutorizado = EnumSiNo.No Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
            Err.Raise 1000
        End If
        Avance "Obteniendo la validación pertinente: Fecha Publicación"
        If Not IsDate(.FechaPublicacion) Then
            p_Error = "La Edición no aparece como publicada"
            Err.Raise 1000
        End If
        Avance "Obteniendo la validación pertinente: Próxima Edición es la activa"
        If .EsUltimaEdicionCalculado <> EnumSiNo.Sí Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            p_Error = "Ha de ser la última Edición"
            Err.Raise 1000
        End If
        Avance "Obteniendo la validación pertinente: Buscando Informe"
        p_URLInforme = .URLArchivoInforme
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Avance "Obteniendo la validación pertinente: Ver si es accesible"
        If fso.FileExists(p_URLInforme) Then
            If FicheroAbierto(p_URLInforme) Then
                p_Error = "Existía un archivo con el mismo nombre que tendrá este, y está abierto"
                Err.Raise 1000
            End If
        End If
        
        
    End With
    
    ValidacionEliminarPublicacion = "Publicación eliminable"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Publicacion.ValidacionEliminarPublicacion ha producido el error : " & vbNewLine & Err.Description
    End If
End Function






Public Property Get EsUltimaEdicion() As String

    EsUltimaEdicion = m_sEsUltimaEdicion

End Property

Public Property Let EsUltimaEdicion(ByVal sNewValue As String)

    m_sEsUltimaEdicion = sNewValue

End Property

Public Property Get AlgunRiesgoNoCerrado() As EnumSiNo

    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_Estado As EnumRiesgoEstado
    
    On Error GoTo errores
    Me.Error = ""
    
    Set Me.colRiesgos = Nothing
    If Me.colRiesgos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eAlgunRiesgoNoCerrado = EnumSiNo.No
        AlgunRiesgoNoCerrado = m_eAlgunRiesgoNoCerrado
        Exit Property
    End If
    For Each m_IdRiesgo In Me.colRiesgos.keys
        Set m_ObjRiesgo = Me.colRiesgos(m_IdRiesgo)
        With m_ObjRiesgo
            m_Estado = .ESTADOCalculado
            
            If m_Estado <> EnumRiesgoEstado.Retirado And _
                m_Estado <> EnumRiesgoEstado.Aceptado Then
                m_eAlgunRiesgoNoCerrado = EnumSiNo.Sí
                AlgunRiesgoNoCerrado = m_eAlgunRiesgoNoCerrado
                Exit Property
            End If
        End With
        
        
        Set m_ObjRiesgo = Nothing
    Next
    
    m_eAlgunRiesgoNoCerrado = EnumSiNo.No
    AlgunRiesgoNoCerrado = m_eAlgunRiesgoNoCerrado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.AlgunRiesgoNoCerrado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let AlgunRiesgoNoCerrado(ByVal eNewValue As EnumSiNo)

    m_eAlgunRiesgoNoCerrado = eNewValue

End Property

Public Property Get AlgunRiesgoNoDeBiblioteca() As EnumSiNo

    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_RiesgoParaRetipificar As EnumSiNo
    On Error GoTo errores
    Me.Error = ""
    
    If Me.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
        m_eAlgunRiesgoNoDeBiblioteca = EnumSiNo.No
        AlgunRiesgoNoDeBiblioteca = m_eAlgunRiesgoNoDeBiblioteca
        Exit Property
    End If
    Set Me.colRiesgos = Nothing
    If Me.colRiesgos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eAlgunRiesgoNoDeBiblioteca = EnumSiNo.No
        AlgunRiesgoNoDeBiblioteca = m_eAlgunRiesgoNoDeBiblioteca
        Exit Property
    End If
    For Each m_IdRiesgo In Me.colRiesgos.keys
        Set m_ObjRiesgo = Me.colRiesgos(m_IdRiesgo)
        m_RiesgoParaRetipificar = m_ObjRiesgo.RiesgoParaRetipificar
        m_Error = m_ObjRiesgo.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_RiesgoParaRetipificar = EnumSiNo.Sí Then
            AlgunRiesgoNoDeBiblioteca = m_RiesgoParaRetipificar
            Exit Property
        End If
        
        
        Set m_ObjRiesgo = Nothing
    Next
    
    m_eAlgunRiesgoNoDeBiblioteca = EnumSiNo.No
    AlgunRiesgoNoDeBiblioteca = m_eAlgunRiesgoNoDeBiblioteca
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.AlgunRiesgoNoDeBiblioteca ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let AlgunRiesgoNoDeBiblioteca(ByVal eNewValue As EnumSiNo)

    m_eAlgunRiesgoNoDeBiblioteca = eNewValue

End Property


Public Function DescripcionYCausaRaizRepetida( _
                                                p_Descripcion As String, _
                                                p_CausaRaiz As String, _
                                                Optional p_IDRiesgoAExcluir As String, _
                                                Optional ByRef p_Error As String _
                                                ) As EnumSiNo
                               
    Dim m_ID As Variant
    Dim m_ObjRiesgo As riesgo
    
    Dim m_Proyecto As Proyecto
    
    On Error GoTo errores
    
    
    
    If Me.colRiesgos Is Nothing Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        DescripcionYCausaRaizRepetida = EnumSiNo.No
        Exit Function
    End If
    
    For Each m_ID In Me.colRiesgos.keys
        If Nz(m_ID, "") <> "" Then
            Set m_ObjRiesgo = Me.colRiesgos(m_ID)
            If p_IDRiesgoAExcluir <> "" Then
                If p_IDRiesgoAExcluir = m_ObjRiesgo.IDRiesgo Then
                    GoTo siguiente
                End If
            End If
            With m_ObjRiesgo
                If .Descripcion = p_Descripcion Then
                    If .CausaRaiz = p_CausaRaiz Then
                        Set m_ObjRiesgo = Nothing
                        DescripcionYCausaRaizRepetida = EnumSiNo.Sí
                        Exit Function
                    End If
                End If
            End With
            
            

            Set m_ObjRiesgo = Nothing
        End If
siguiente:
    Next
    DescripcionYCausaRaizRepetida = EnumSiNo.No
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DescripcionYCausaRaizRepetida ha devuelto el error: " & Err.Description
    End If
End Function

Public Function DescripcionYCausaRaizRepetidaDeRiesgosNoTrasladados( _
                                                                    p_Descripcion As String, _
                                                                    p_CausaRaiz As String, _
                                                                    Optional ByRef p_Error As String _
                                                                    ) As EnumSiNo
                               
    Dim m_ID As Variant
    Dim m_ObjRiesgoExt As RiesgoExterno
    
    
    
    On Error GoTo errores
    
    
    If Me.ColRiesgosExternos Is Nothing Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        DescripcionYCausaRaizRepetidaDeRiesgosNoTrasladados = EnumSiNo.No
        Exit Function
    End If
    
    For Each m_ID In Me.ColRiesgosExternos.keys
        If Nz(m_ID, "") <> "" Then
            Set m_ObjRiesgoExt = Me.ColRiesgosExternos(m_ID)
            
            With m_ObjRiesgoExt
                If .CodRiesgo = "" Then
                    If .Descripcion = p_Descripcion Then
                        If .CausaRaiz = p_CausaRaiz Then
                            Set m_ObjRiesgoExt = Nothing
                            DescripcionYCausaRaizRepetidaDeRiesgosNoTrasladados = EnumSiNo.Sí
                            Exit Function
                        End If
                    End If
                End If
                
            End With
            
            

            Set m_ObjRiesgoExt = Nothing
        End If

    Next
    DescripcionYCausaRaizRepetidaDeRiesgosNoTrasladados = EnumSiNo.No
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DescripcionYCausaRaizRepetidaDeRiesgosNoTrasladados ha devuelto el error: " & Err.Description
    End If
End Function



Public Property Get ColLogPulicaciones() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColLogPulicaciones Is Nothing Then
        Set ColLogPulicaciones = m_objColLogPulicaciones
        Exit Property
    End If
    
    Set m_objColLogPulicaciones = Constructor.getLogPublicaciones(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColLogPulicaciones = m_objColLogPulicaciones
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColLogPulicaciones ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColLogPulicaciones(ByVal objNewValue As Variant)

    Set m_objColLogPulicaciones = objNewValue

End Property
Public Property Get ColLogPublicacionesProyecto() As Scripting.Dictionary
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColLogPublicacionesProyecto Is Nothing Then
        Set ColLogPublicacionesProyecto = m_objColLogPublicacionesProyecto
        Exit Property
    End If

    Set m_objColLogPublicacionesProyecto = Constructor.getLogPublicacionesProyecto(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColLogPublicacionesProyecto = m_objColLogPublicacionesProyecto
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColLogPublicacionesProyecto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColLogPublicacionesProyecto(ByVal objNewValue As Variant)

    Set m_objColLogPublicacionesProyecto = objNewValue

End Property
Public Property Get ColLogCorreosResponsables() As Scripting.Dictionary
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColLogCorreosResponsables Is Nothing Then
        Set ColLogCorreosResponsables = m_objColLogCorreosResponsables
        Exit Property
    End If
    
    Set m_objColLogCorreosResponsables = Constructor.getLogCorreosResponsablesProyecto(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColLogCorreosResponsables = m_objColLogCorreosResponsables
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColLogCorreosResponsables ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColLogCorreosResponsables(ByVal objNewValue As Variant)

    Set m_objColLogCorreosResponsables = objNewValue

End Property

Public Property Get ColSuministradoresEvidencias() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColSuministradoresEvidencias Is Nothing Then
        Set ColSuministradoresEvidencias = m_objColSuministradoresEvidencias
        Exit Property
    End If
    Set m_objColSuministradoresEvidencias = SuministradoresHelper.getSuministradoresEvidenciasEnEdicion(p_IDEdicion:=Me.IDEdicion, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
   
    Set ColSuministradoresEvidencias = m_objColSuministradoresEvidencias
    
     Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColSuministradoresEvidencias ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Sub ColSuministradoresEvidenciasVaciar()
    Set m_objColSuministradoresEvidencias = Nothing
End Sub

Public Property Get ColSuministradores() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColSuministradores Is Nothing Then
        Set ColSuministradores = m_objColSuministradores
        Exit Property
    End If
    Set m_objColSuministradores = SuministradoresHelper.getSuministradoresEnEdicion(p_IDEdicion:=Me.IDEdicion, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
   
    Set ColSuministradores = m_objColSuministradores
    
     Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColSuministradores ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Sub ColSuministradoresVaciar()
    Set m_objColSuministradores = Nothing
End Sub
Public Property Get ColSuministradoresFaltan() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColSuministradoresFaltan Is Nothing Then
        Set ColSuministradoresFaltan = m_ObjColSuministradoresFaltan
        Exit Property
    End If
    If Me.Proyecto Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    Set m_ObjColSuministradoresFaltan = SuministradoresHelper.getSuministradoresFaltantesEnEdicion( _
                                            p_IDExpediente:=Me.Proyecto.IDExpediente, _
                                            p_IDEdicion:=Me.IDEdicion, _
                                            p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColSuministradoresFaltan = m_ObjColSuministradoresFaltan
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColSuministradoresFaltan ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Sub ColSuministradoresFaltanVaciar()
    Set m_ObjColSuministradoresFaltan = Nothing
End Sub
Public Property Get ColSuministradoresSobran() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjColSuministradoresSobran Is Nothing Then
        Set ColSuministradoresSobran = m_ObjColSuministradoresSobran
        Exit Property
    End If
    If Me.Proyecto Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    Set m_ObjColSuministradoresSobran = SuministradoresHelper.getSuministradoresSobrantesEnEdicion( _
                                            p_IDExpediente:=Me.Proyecto.IDExpediente, _
                                            p_IDEdicion:=Me.IDEdicion, _
                                            p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColSuministradoresSobran = m_ObjColSuministradoresSobran
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColSuministradoresSobran ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Sub ColSuministradoresSobranVaciar()
    Set m_ObjColSuministradoresSobran = Nothing
End Sub
Public Property Get Publicable() As EnumSiNo
    Dim m_ErrorLocal As String
    Dim m_PublicableBool As Boolean
    Dim m_UsarCache As String

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""

    m_UsarCache = "No"
    On Error Resume Next
    m_UsarCache = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    On Error GoTo errores

    If UCase$(Trim$(m_UsarCache)) = "NO" Then
        m_ErrorLocal = ""
        If CalcularPublicabilidadEdicion(Me, m_PublicableBool, , m_ErrorLocal) = EnumSiNo.No Then
            m_Error = m_ErrorLocal
            If m_Error = "" Then
                m_Error = "No se ha podido establecer la publicabilidad"
            End If
            Err.Raise 1000
        End If
    Else
        m_ErrorLocal = ""
        If CachePublicabilidad_LeerEdicionPublicable(Me.IDEdicion, m_PublicableBool, m_ErrorLocal) = EnumSiNo.No Then
            If m_ErrorLocal <> "" Then
                m_Error = m_ErrorLocal
                Err.Raise 1000
            End If
            m_ErrorLocal = ""
            If CachePublicabilidad_RecalcularEdicion(Me, Nothing, m_ErrorLocal) = EnumSiNo.No Then
                If m_ErrorLocal <> "" Then
                    m_Error = m_ErrorLocal
                    Err.Raise 1000
                End If
            End If
            m_ErrorLocal = ""
            If CachePublicabilidad_LeerEdicionPublicable(Me.IDEdicion, m_PublicableBool, m_ErrorLocal) = EnumSiNo.No Then
                If m_ErrorLocal <> "" Then
                    m_Error = m_ErrorLocal
                    Err.Raise 1000
                End If
                m_ErrorLocal = ""
                If CalcularPublicabilidadEdicion(Me, m_PublicableBool, , m_ErrorLocal) = EnumSiNo.No Then
                    m_Error = m_ErrorLocal
                    If m_Error = "" Then
                        m_Error = "No se ha podido establecer la publicabilidad"
                    End If
                    Err.Raise 1000
                End If
            End If
        End If
    End If
    m_ePublicable = IIf(m_PublicableBool, EnumSiNo.Sí, EnumSiNo.No)
    Publicable = m_ePublicable
    Exit Property

errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Publicación.Publicable ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Let Publicable(m_valor As EnumSiNo)
    m_ePublicable = m_valor

End Property
Public Property Get EvidenciasSuministradoresRequeridas() As EnumSiNo
    
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    Dim m_ColSuministradores As Scripting.Dictionary

    Me.Proyecto.ColSuministradoresVaciar
    Set m_ColSuministradores = Me.Proyecto.ColSuministradores

    ' SuministradoresHelper siempre devuelve Dictionary inicializado; vacío => Count=0.
    If Not m_ColSuministradores Is Nothing And m_ColSuministradores.Count > 0 Then
        m_eEvidenciasSuministradoresRequeridas = EnumSiNo.Sí
    Else
        m_eEvidenciasSuministradoresRequeridas = EnumSiNo.No
    End If
    EvidenciasSuministradoresRequeridas = m_eEvidenciasSuministradoresRequeridas
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.EvidenciasSuministradoresRequeridas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EvidenciasSuministradoresCompletadas() As EnumSiNo
    
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    m_eEvidenciasSuministradoresCompletadas = SuministradoresHelper.EvidenciasSuministradoresCompletadas(Me.IDEdicion, EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    EvidenciasSuministradoresCompletadas = m_eEvidenciasSuministradoresCompletadas
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.EvidenciasSuministradoresCompletadas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EdicionPrimeraYActivaConPermisoUsuario() As EnumSiNo
    
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    m_eEdicionPrimeraYActivaConPermisoUsuario = getEsEdicionPrimeraYActivaConPermisoUsuario(Me, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EdicionPrimeraYActivaConPermisoUsuario = m_eEdicionPrimeraYActivaConPermisoUsuario
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.EdicionPrimeraYActivaConPermisoUsuario ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function BorrarCambiosEdicion( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_SQL As String
    Dim m_IDProyecto As String
    Dim m_Edicion As String
    Dim m_EdicionInicial As String
    
    On Error GoTo errores
    m_IDProyecto = Me.IDProyecto
    m_Edicion = Me.Edicion
    If Not Me.EdicionAnterior Is Nothing Then
        m_EdicionInicial = Me.EdicionAnterior.Edicion
    End If
    
    If Not IsNumeric(m_EdicionInicial) Then
        m_SQL = "DELETE * FROM TbCambios " & _
            "WHERE IDProyecto=" & m_IDProyecto & " AND EdicionFinal=" & m_Edicion & ";"
    Else
        m_SQL = "DELETE * FROM TbCambios " & _
            "WHERE IDProyecto=" & m_IDProyecto & " AND EdicionFinal=" & m_Edicion & " AND EdicionInicial=" & m_EdicionInicial & ";"
    End If
    getdb().Execute m_SQL
    Set Me.colCambiosConEdicionAnterior = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método BorrarCambiosEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function getColCambiosEntreDosObjetos( _
                                            p_Obj As Object, _
                                            Optional p_ObjAnterior As Object, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    
    
    'getColCambiosEntreDosObjetos.Add m_NombreCampo, m_ObjCambio
    'm_ObjCambio--->m_NombreCampo,  m_ValorAnterior  m_ValorActual
    On Error GoTo errores
    
    If p_Obj Is Nothing Then
        p_Error = "No se han indicado los datos actuales"
        Err.Raise 1000
    End If
    If p_ObjAnterior Is Nothing Then
        
        Set p_ObjAnterior = getObjetoAnterior(p_Obj, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Set getColCambiosEntreDosObjetos = getColCambios(p_Obj, p_ObjAnterior, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColCambiosEntreDosObjetos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Private Function getColCambios( _
                                p_ObjAlFinal As Object, _
                                p_ObjAlInicio As Object, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    
    'getColCambios.Add m_NombreCampo, m_ValorActual & "|" & m_ValorAnterior
    Dim m_NombreCampo As Variant
    Dim m_ColCampos As Collection
    Dim m_ValorActual As String
    Dim m_ValorAnterior As String
    Dim m_objColCampos As Collection
    
    Dim m_ConsiderarCampo As EnumSiNo
    On Error GoTo errores
    
    If p_ObjAlFinal Is Nothing Then
        p_Error = "No se han indicado los datos actuales"
        Err.Raise 1000
    End If
    If p_ObjAlInicio Is Nothing Then
        p_Error = "No se han indicado los datos iniciales"
        Err.Raise 1000
    End If
    Set m_objColCampos = p_ObjAlFinal.ColCamposParaCambios
    p_Error = p_ObjAlFinal.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For Each m_NombreCampo In m_objColCampos
        m_ValorActual = p_ObjAlFinal.getPropiedad(m_NombreCampo, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ValorActual = "" Then m_ValorActual = "NULL"
        
        m_ValorAnterior = p_ObjAlInicio.getPropiedad(m_NombreCampo, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ValorAnterior = "" Then m_ValorAnterior = "NULL"
        
        If m_ValorActual <> m_ValorAnterior Then
            If getColCambios Is Nothing Then
                Set getColCambios = New Scripting.Dictionary
                getColCambios.CompareMode = TextCompare
            End If
            If Not getColCambios.Exists(m_NombreCampo) Then
                getColCambios.Add m_NombreCampo, m_ValorActual & "|" & m_ValorAnterior
            End If
            
        End If
siguiente:
        
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function FechaMaxProximaPublicacionGrabar( _
                                                    Optional ByRef p_Error As String _
                                                     ) As String
    
    Dim m_SQL As String
    Dim m_FechaMax
    
    On Error GoTo errores
    m_FechaMax = Me.FechaMaxProximaPublicacionCalculado
    If Not IsDate(m_FechaMax) Then
        FechaMaxProximaPublicacionGrabar = m_FechaMax
        Exit Function
    End If
    
    m_SQL = "UPDATE TbProyectosEdiciones SET FechaMaxProximaPublicacion = #" & Format(m_FechaMax, "mm/dd/yyyy") & "# " & _
            "WHERE IDEdicion=" & Me.IDEdicion & ";"
    getdb.Execute m_SQL
    FechaMaxProximaPublicacionGrabar = m_FechaMax
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FechaMaxProximaPublicacionGrabar ha devuelto el error: " & Err.Description
    End If
End Function



Public Function GrabarCambiosEnEdicion( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
   
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_ObjRiesgoAnterior As riesgo
    Dim m_ObjColRiesgosEdicion As Scripting.Dictionary
    Dim m_ObjColRiesgosEdicionAnterior As Scripting.Dictionary
    Dim m_CodRiesgo As String
    Dim m_CodRiesgoAnterior As String
    Dim m_NombreCampo As Variant
    
    Dim m_ObjColPM As Scripting.Dictionary
    Dim m_ObjColPMAnterior As Scripting.Dictionary
    Dim m_ObjPM As PM
    Dim m_IDPM As Variant
    Dim m_ObjPMAnterior As PM
    Dim m_ObjColPMA As Scripting.Dictionary
    Dim m_ObjColPMAAnterior As Scripting.Dictionary
    Dim m_ObjPMA As PMAccion
    Dim m_ObjPMAAnterior As PMAccion
    Dim m_IDPMA As Variant
    Dim m_ObjColPC As Scripting.Dictionary
    Dim m_ObjColPCAnterior As Scripting.Dictionary
    Dim m_ObjPC As PC
    Dim m_IDPC As Variant
    Dim m_ObjPCAnterior As PC
    Dim m_ObjPCA As PCAccion
    Dim m_ObjPCAAnterior As PCAccion
    Dim m_IDPCA As Variant
    Dim m_ObjColPCA As Scripting.Dictionary
    Dim m_ObjColPCAAnterior As Scripting.Dictionary
    Dim m_objCambios As Scripting.Dictionary
    Dim m_valorFinal As String
    Dim m_valorInicial As String
    Dim m_ObjColCambios As Scripting.Dictionary
    Dim m_ObjCambio As Cambio
    Dim dato
    Dim m_Resultado As String
    Dim m_EdicionInicial As String
    Dim m_EDicionFinal As String
    Dim m_ObjEdicionAnterior As Edicion
    On Error GoTo errores
    
    Set m_ObjEdicionAnterior = Me.EdicionAnterior
    If Not m_ObjEdicionAnterior Is Nothing Then
        m_EdicionInicial = m_ObjEdicionAnterior.Edicion
    End If
    BorrarCambiosEdicion p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_EDicionFinal = Me.Edicion
    
    If m_ObjEdicionAnterior Is Nothing Then
        Set m_ObjCambio = New Cambio
        With m_ObjCambio
            .IDProyecto = Me.IDProyecto
            
            .EdicionFinal = "1"
            .riesgo = "Primera Edición"
            
            .Registrar p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End With
        Set m_ObjCambio = Nothing
        Exit Function
    End If
    
    Set Me.colRiesgos = Nothing
    Set m_ObjColRiesgosEdicion = Me.colRiesgos
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Avance "Registrando Cambios....Edición "
    'cambios previos a los de los riesgos (Elaborado, revisado, aprobado)
    With m_ObjEdicionAnterior
        m_valorFinal = Me.Aprobado
        m_valorInicial = .Aprobado
        If m_valorInicial <> m_valorFinal Then
            Set m_ObjCambio = New Cambio
            With m_ObjCambio
                .IDProyecto = Me.IDProyecto
                .EdicionInicial = m_EdicionInicial
                .EdicionFinal = m_EDicionFinal
                .riesgo = "N/A"
                .NombreCampo = "UsuarioAprueba"
                .Descripcion = "R " & m_NombreCampo & ":" & m_valorFinal & " " & m_NombreCampo & ":" & m_valorInicial
                .Registrar p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End With
            Set m_ObjCambio = Nothing
        End If
        m_valorFinal = Me.Elaborado
        m_valorInicial = .Elaborado
        If m_valorInicial <> m_valorFinal Then
            Set m_ObjCambio = New Cambio
            With m_ObjCambio
                .IDProyecto = Me.IDProyecto
                .EdicionInicial = m_EdicionInicial
                .EdicionFinal = m_EDicionFinal
                .riesgo = "N/A"
                .NombreCampo = "UsuarioElabora"
                .Descripcion = "R " & m_NombreCampo & ":" & m_valorFinal & " " & m_NombreCampo & ":" & m_valorInicial
                .Registrar p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End With
            Set m_ObjCambio = Nothing
        End If
        m_valorFinal = Me.Revisado
        m_valorInicial = .Revisado
        If m_valorInicial <> m_valorFinal Then
            Set m_ObjCambio = New Cambio
            With m_ObjCambio
                .IDProyecto = Me.IDProyecto
                .EdicionInicial = m_EdicionInicial
                .EdicionFinal = m_EDicionFinal
                .riesgo = "N/A"
                .NombreCampo = "UsuarioRevisa"
                .Descripcion = "R " & m_NombreCampo & ":" & m_valorFinal & " " & m_NombreCampo & ":" & m_valorInicial
                .Registrar p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End With
            Set m_ObjCambio = Nothing
        End If
    End With
    
    
    If m_ObjColRiesgosEdicion Is Nothing Then
       Exit Function
    End If
    
    Set m_ObjEdicionAnterior.colRiesgos = Nothing
    Set m_ObjColRiesgosEdicionAnterior = m_ObjEdicionAnterior.colRiesgos
    p_Error = m_ObjEdicionAnterior.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjColRiesgosEdicionAnterior Is Nothing Then
        Exit Function
    End If
    
    
    
    'PARTO DE LOS RIESGOS DE LA Edición MÁS NUEVA (los riesgos no se pueden borrar)
    If Not m_ObjColRiesgosEdicion Is Nothing Then
        For Each m_IdRiesgo In m_ObjColRiesgosEdicion.keys
            If Nz(m_IdRiesgo, "") = "" Then
                GoTo siguienteRiesgo
            End If
            
            Set m_ObjRiesgo = m_ObjColRiesgosEdicion(m_IdRiesgo)
            'If m_ObjRiesgo.CodigoRiesgo = "R003" Then Stop
            Avance "Registrando Cambios....Edición Riesgo " & m_ObjRiesgo.CodigoRiesgo
            Set m_ObjRiesgoAnterior = m_ObjRiesgo.RiesgoEdicionAnterior
            p_Error = m_ObjRiesgo.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjRiesgoAnterior Is Nothing Then
                'no existía El Riesgo en la anterior Edición
                
                Set m_ObjCambio = New Cambio
                With m_ObjCambio
                    .IDProyecto = Me.IDProyecto
                    .EdicionInicial = m_EdicionInicial
                    .EdicionFinal = m_EDicionFinal
                    .riesgo = m_ObjRiesgo.CodigoRiesgo
                    .Descripcion = "Alta del Riesgo"
                    .Registrar p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                End With
                Set m_ObjCambio = Nothing
                GoTo siguienteRiesgo
            End If
            'getColCambiosEntreDosObjetos.Add m_NombreCampo, m_ObjCambio
            'm_ObjCambio--->m_NombreCampo,  m_ValorAnterior  m_ValorActual
            
            Set m_objCambios = Nothing
            Set m_objCambios = getColCambiosEntreDosObjetos(m_ObjRiesgo, m_ObjRiesgoAnterior, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not m_objCambios Is Nothing Then
                For Each m_NombreCampo In m_objCambios
                    'm_Resultado=m_ValorActual & "|" & m_ValorAnterior
                    m_Resultado = m_objCambios(m_NombreCampo)
                    dato = Split(m_Resultado, "|")
                    m_valorFinal = dato(0)
                    m_valorInicial = dato(1)
                    Set m_ObjCambio = New Cambio
                    With m_ObjCambio
                        .IDProyecto = Me.IDProyecto
                        .EdicionInicial = m_EdicionInicial
                        .EdicionFinal = m_EDicionFinal
                        .riesgo = m_ObjRiesgo.CodigoRiesgo
                        .NombreCampo = CStr(m_NombreCampo)
                        .Descripcion = "R " & m_NombreCampo & ":" & m_valorFinal & " " & m_NombreCampo & ":" & m_valorInicial
                        .Registrar p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    End With
                    Set m_ObjCambio = Nothing
                Next
            End If
            Set m_ObjColPM = m_ObjRiesgo.ColPMs
            p_Error = m_ObjRiesgo.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            '-------------------------------
            ' PLANES DE MITIGACIÓN
            '-------------------------------
            Set m_ObjColPMAnterior = m_ObjRiesgoAnterior.ColPMs
            p_Error = m_ObjRiesgoAnterior.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Not m_ObjColPM Is Nothing Then
                If Not m_ObjColPMAnterior Is Nothing Then
                    For Each m_IDPM In m_ObjColPM.keys
                        Set m_ObjPM = m_ObjColPM(m_IDPM)
                        Set m_ObjPMAnterior = m_ObjPM.PMEdicionAnterior
                        p_Error = m_ObjPM.Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        If m_ObjPMAnterior Is Nothing Then
                            Set m_ObjCambio = New Cambio
                            With m_ObjCambio
                                .IDProyecto = Me.IDProyecto
                                .EdicionInicial = m_EdicionInicial
                                .EdicionFinal = m_EDicionFinal
                                .riesgo = m_ObjRiesgo.CodigoRiesgo
                                .Descripcion = "PM " & " " & "Alta de Plan de Mitigación: " & m_ObjPM.CodMitigacion
                                .Registrar p_Error
                                If p_Error <> "" Then
                                    Err.Raise 1000
                                End If
                            End With
                        Else
                            Set m_objCambios = Nothing
                            Set m_objCambios = getColCambiosEntreDosObjetos(m_ObjPM, m_ObjPMAnterior, p_Error)
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                            If Not m_objCambios Is Nothing Then
                                For Each m_NombreCampo In m_objCambios.keys
                                    'm_Resultado=m_ValorActual & "|" & m_ValorAnterior
                                    m_Resultado = m_objCambios(m_NombreCampo)
                                    dato = Split(m_Resultado, "|")
                                    m_valorFinal = dato(0)
                                    m_valorInicial = dato(1)
                                    Set m_ObjCambio = New Cambio
                                    With m_ObjCambio
                                        .IDProyecto = Me.IDProyecto
                                        .EdicionInicial = m_EdicionInicial
                                        .EdicionFinal = m_EDicionFinal
                                        .riesgo = m_ObjRiesgo.CodigoRiesgo
                                        .NombreCampo = CStr(m_NombreCampo)
                                        .Descripcion = "PM " & m_NombreCampo & ":" & m_valorFinal & " " & m_NombreCampo & ":" & m_valorInicial
                                        .Registrar p_Error
                                        If p_Error <> "" Then
                                            Err.Raise 1000
                                        End If
                                    End With
                                    Set m_ObjCambio = Nothing
                                                    
                                Next
                            End If
                            'aquí comparamos las acciones de cada plan idéntico
                            Set m_ObjColPMA = m_ObjPM.colAcciones
                            p_Error = m_ObjPM.Error
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                            Set m_ObjColPMAAnterior = m_ObjPMAnterior.colAcciones
                            p_Error = m_ObjPMAnterior.Error
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                            If Not m_ObjColPMA Is Nothing Then
                                If Not m_ObjColPMAAnterior Is Nothing Then
                                    For Each m_IDPMA In m_ObjColPMA.keys
                                        Set m_ObjPMA = m_ObjColPMA(m_IDPMA)
                                        Set m_ObjPMAAnterior = m_ObjPMA.PMAccionEdicionAnterior
                                        p_Error = m_ObjPMA.Error
                                        If p_Error <> "" Then
                                            Err.Raise 1000
                                        End If
                                        If m_ObjPMAAnterior Is Nothing Then
                                            Set m_ObjCambio = New Cambio
                                            With m_ObjCambio
                                                .IDProyecto = Me.IDProyecto
                                                .EdicionInicial = m_EdicionInicial
                                                .EdicionFinal = m_EDicionFinal
                                                .riesgo = m_ObjRiesgo.CodigoRiesgo
                                                .Descripcion = "PMA " & " " & "Alta de Acción: " & m_ObjPMA.CodAccion & " del Plan de Mitigación: " & m_ObjPMA.Mitigacion.CodMitigacion
                                                .Registrar p_Error
                                                If p_Error <> "" Then
                                                    Err.Raise 1000
                                                End If
                                            End With
                                            Set m_ObjCambio = Nothing
                                            
                                            
                                        Else
                                            Set m_objCambios = Nothing
                                            Set m_objCambios = getColCambiosEntreDosObjetos(m_ObjPMA, m_ObjPMAAnterior, p_Error)
                                            If p_Error <> "" Then
                                                Err.Raise 1000
                                            End If
                                            If Not m_objCambios Is Nothing Then
                                                For Each m_NombreCampo In m_objCambios.keys
                                                    'm_Resultado=m_ValorActual & "|" & m_ValorAnterior
                                                    m_Resultado = m_objCambios(m_NombreCampo)
                                                    dato = Split(m_Resultado, "|")
                                                    m_valorFinal = dato(0)
                                                    m_valorInicial = dato(1)
                                                    Set m_ObjCambio = New Cambio
                                                    With m_ObjCambio
                                                        .IDProyecto = Me.IDProyecto
                                                        .EdicionInicial = m_EdicionInicial
                                                        .EdicionFinal = m_EDicionFinal
                                                        .riesgo = m_ObjRiesgo.CodigoRiesgo
                                                        .NombreCampo = CStr(m_NombreCampo)
                                                        .Descripcion = "PMA " & m_NombreCampo & ":" & m_valorFinal & " " & m_NombreCampo & ":" & m_valorInicial
                                                        .Registrar p_Error
                                                        If p_Error <> "" Then
                                                            Err.Raise 1000
                                                        End If
                                                    End With
                                                    Set m_ObjCambio = Nothing
                                                    
                                                Next
                                            End If
                                        End If
                                        
                                        Set m_ObjPMA = Nothing
                                        Set m_ObjPMAAnterior = Nothing
                                    Next
                                    For Each m_IDPMA In m_ObjColPMAAnterior.keys
                                        Set m_ObjPMAAnterior = m_ObjColPMAAnterior(m_IDPMA)
                                        Set m_ObjPMA = m_ObjPMAAnterior.PMAccionEdicionSiguiente
                                        p_Error = m_ObjPMAAnterior.Error
                                        If p_Error <> "" Then
                                            Err.Raise 1000
                                        End If
                                        If m_ObjPMA Is Nothing Then
                                            Set m_ObjCambio = New Cambio
                                            With m_ObjCambio
                                                .IDProyecto = Me.IDProyecto
                                                .EdicionInicial = m_EdicionInicial
                                                .EdicionFinal = m_EDicionFinal
                                                .riesgo = m_ObjRiesgo.CodigoRiesgo
                                                .Descripcion = "PMA " & " " & "Baja de Acción: " & m_ObjPMAAnterior.CodAccion & " del Plan de Mitigación: " & m_ObjPMAAnterior.Mitigacion.CodMitigacion
                                                .Registrar p_Error
                                                If p_Error <> "" Then
                                                    Err.Raise 1000
                                                End If
                                            End With
                                            Set m_ObjCambio = Nothing
                                           
                                        End If
                                        Set m_ObjPMAAnterior = Nothing
                                        Set m_ObjPMA = Nothing
                                    Next
                                Else
                                    'm_ObjColPMAAnterior Is Nothing
                                    For Each m_IDPMA In m_ObjColPMA.keys
                                        Set m_ObjPMA = m_ObjColPMA(m_IDPMA)
                                        Set m_ObjCambio = New Cambio
                                        With m_ObjCambio
                                            .IDProyecto = Me.IDProyecto
                                            .EdicionInicial = m_EdicionInicial
                                            .EdicionFinal = m_EDicionFinal
                                            .riesgo = m_ObjRiesgo.CodigoRiesgo
                                            .Descripcion = "PMA " & " " & "Alta de Acción: " & m_ObjPMA.CodAccion & " del Plan de Mitigación: " & m_ObjPMA.Mitigacion.CodMitigacion
                                            .Registrar p_Error
                                            If p_Error <> "" Then
                                                Err.Raise 1000
                                            End If
                                        End With
                                        Set m_ObjCambio = Nothing
                                        
                                        Set m_ObjPMA = Nothing
                                    Next
                                End If
                            Else
                                If Not m_ObjColPMAAnterior Is Nothing Then
                                    For Each m_IDPMA In m_ObjColPMAAnterior.keys
                                        Set m_ObjPMAAnterior = m_ObjColPMAAnterior(m_IDPMA)
                                        If Not m_ObjColPMA Is Nothing Then
                                            
                                        
                                        Set m_ObjPMA = m_ObjColPMA(m_IDPMA)
                                        Set m_ObjCambio = New Cambio
                                        With m_ObjCambio
                                            .IDProyecto = Me.IDProyecto
                                            .EdicionInicial = m_EdicionInicial
                                            .EdicionFinal = m_EDicionFinal
                                            .riesgo = m_ObjRiesgo.CodigoRiesgo
                                            .Descripcion = "PM" & " " & "Baja de Acción: " & m_ObjPMAAnterior.CodAccion & " del Plan de Mitigación: " & m_ObjPMAAnterior.Mitigacion.CodMitigacion
                                            .Registrar p_Error
                                            If p_Error <> "" Then
                                                Err.Raise 1000
                                            End If
                                        End With
                                        Set m_ObjCambio = Nothing
                                        End If
                                        
                                        Set m_ObjPMAAnterior = Nothing
                                    Next
                                End If
                            End If
                            
                        End If
                        Set m_ObjPM = Nothing
                        Set m_ObjPMAnterior = Nothing
                    Next
                    For Each m_IDPM In m_ObjColPMAnterior.keys
                        Set m_ObjPMAnterior = m_ObjColPMAnterior(m_IDPM)
                        Set m_ObjPM = m_ObjPMAnterior.PMEdicionSiguiente
                        p_Error = m_ObjPMAnterior.Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        If m_ObjPM Is Nothing Then
                            Set m_ObjCambio = New Cambio
                            With m_ObjCambio
                                .IDProyecto = Me.IDProyecto
                                .EdicionInicial = m_EdicionInicial
                                .EdicionFinal = m_EDicionFinal
                                .riesgo = m_ObjRiesgo.CodigoRiesgo
                                .Descripcion = "PM " & " " & "Baja de Plan de Mitigación: " & m_ObjPMAnterior.CodMitigacion
                                .Registrar p_Error
                                If p_Error <> "" Then
                                    Err.Raise 1000
                                End If
                            End With
                            Set m_ObjCambio = Nothing
                            
                            
                        End If
                    Next
                Else
                    'm_ObjColPMAnterior-->Nulo
                    For Each m_IDPM In m_ObjColPM.keys
                        Set m_ObjPM = m_ObjColPM(m_IDPM)
                        Set m_ObjCambio = New Cambio
                        With m_ObjCambio
                            .IDProyecto = Me.IDProyecto
                            .EdicionInicial = m_EdicionInicial
                            .EdicionFinal = m_EDicionFinal
                            .riesgo = m_ObjRiesgo.CodigoRiesgo
                            .Descripcion = "PM " & " " & "Alta de Plan de Mitigación: " & m_ObjPM.CodMitigacion
                            .Registrar p_Error
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                        End With
                        Set m_ObjCambio = Nothing
                    Next
                End If
            Else
                If Not m_ObjColPMAnterior Is Nothing Then
                    For Each m_IDPM In m_ObjColPMAnterior.keys
                        Set m_ObjPMAnterior = m_ObjColPMAnterior(m_IDPM)
                        Set m_ObjCambio = New Cambio
                        With m_ObjCambio
                            .IDProyecto = Me.IDProyecto
                            .EdicionInicial = m_EdicionInicial
                            .EdicionFinal = m_EDicionFinal
                            .riesgo = m_ObjRiesgo.CodigoRiesgo
                            .Descripcion = "PM " & " " & "Baja de Plan de Mitigación: " & m_ObjPMAnterior.CodMitigacion
                            .Registrar p_Error
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                        End With
                        Set m_ObjCambio = Nothing
                                                
                    Next
                End If
            End If
            '-------------------------------
            ' PLANES DE CONTINGENCIA
            '-------------------------------
            Set m_ObjColPC = m_ObjRiesgo.ColPCs
            p_Error = m_ObjRiesgo.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            Set m_ObjColPCAnterior = m_ObjRiesgoAnterior.ColPCs
            p_Error = m_ObjRiesgoAnterior.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            If Not m_ObjColPC Is Nothing Then
                If Not m_ObjColPCAnterior Is Nothing Then
                    For Each m_IDPC In m_ObjColPC.keys
                        Set m_ObjPC = m_ObjColPC(m_IDPC)
                        Set m_ObjPCAnterior = m_ObjPC.PCEdicionAnterior
                        p_Error = m_ObjPC.Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        If m_ObjPCAnterior Is Nothing Then
                            Set m_ObjCambio = New Cambio
                            With m_ObjCambio
                                .IDProyecto = Me.IDProyecto
                                .EdicionInicial = m_EdicionInicial
                                .EdicionFinal = m_EDicionFinal
                                .Descripcion = "PC " & " " & "Alta de plan de contingencia: " & m_ObjPC.CodContingencia
                                .Registrar p_Error
                                If p_Error <> "" Then
                                    Err.Raise 1000
                                End If
                            End With
                            Set m_ObjCambio = Nothing
                            
                            
                        Else
                            Set m_objCambios = Nothing
                            Set m_objCambios = getColCambiosEntreDosObjetos(m_ObjPC, m_ObjPCAnterior, p_Error)
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                            If Not m_objCambios Is Nothing Then
                                For Each m_NombreCampo In m_objCambios
                                    'm_Resultado=m_ValorActual & "|" & m_ValorAnterior
                                    m_Resultado = m_objCambios(m_NombreCampo)
                                    dato = Split(m_Resultado, "|")
                                    m_valorFinal = dato(0)
                                    m_valorInicial = dato(1)
                                    Set m_ObjCambio = New Cambio
                                    With m_ObjCambio
                                        .IDProyecto = Me.IDProyecto
                                        .EdicionInicial = m_EdicionInicial
                                        .EdicionFinal = m_EDicionFinal
                                        .riesgo = m_ObjRiesgo.CodigoRiesgo
                                        .NombreCampo = CStr(m_NombreCampo)
                                        .Descripcion = "PC" & m_NombreCampo & ":" & m_valorFinal & " " & m_NombreCampo & ":" & m_valorInicial
                                        .Registrar p_Error
                                        If p_Error <> "" Then
                                            Err.Raise 1000
                                        End If
                                    End With
                                    Set m_ObjCambio = Nothing
                                Next
                            End If
                            'aquí comparamos las acciones de cada plan idéntico
                            Set m_ObjColPCA = m_ObjPC.colAcciones
                            p_Error = m_ObjPC.Error
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                            Set m_ObjColPCAAnterior = m_ObjPCAnterior.colAcciones
                            p_Error = m_ObjPCAnterior.Error
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                            If Not m_ObjColPCA Is Nothing Then
                                If Not m_ObjColPCAAnterior Is Nothing Then
                                    For Each m_IDPCA In m_ObjColPCA.keys
                                        Set m_ObjPCA = m_ObjColPCA(m_IDPCA)
                                        Set m_ObjPCAAnterior = m_ObjPCA.PCAccionEdicionAnterior
                                        p_Error = m_ObjPCA.Error
                                        If p_Error <> "" Then
                                            Err.Raise 1000
                                        End If
                                        If m_ObjPCAAnterior Is Nothing Then
                                            Set m_ObjCambio = New Cambio
                                            With m_ObjCambio
                                                .IDProyecto = Me.IDProyecto
                                                .EdicionInicial = m_EdicionInicial
                                                .EdicionFinal = m_EDicionFinal
                                                .riesgo = m_ObjRiesgo.CodigoRiesgo
                                                .Descripcion = "PC" & " " & "Alta de Acción: " & m_ObjPCA.CodAccion & " del plan de contingencia: " & m_ObjPCA.Contingencia.CodContingencia
                                                .Registrar p_Error
                                                If p_Error <> "" Then
                                                    Err.Raise 1000
                                                End If
                                            End With
                                            Set m_ObjCambio = Nothing
                                           
                                        Else
                                            Set m_objCambios = Nothing
                                            Set m_objCambios = getColCambiosEntreDosObjetos(m_ObjPCA, m_ObjPCAAnterior, p_Error)
                                            If p_Error <> "" Then
                                                Err.Raise 1000
                                            End If
                                            If Not m_objCambios Is Nothing Then
                                                For Each m_NombreCampo In m_objCambios.keys
                                                    'm_Resultado=m_ValorActual & "|" & m_ValorAnterior
                                                    m_Resultado = m_objCambios(m_NombreCampo)
                                                    dato = Split(m_Resultado, "|")
                                                    m_valorFinal = dato(0)
                                                    m_valorInicial = dato(1)
                                                    Set m_ObjCambio = New Cambio
                                                    With m_ObjCambio
                                                        .IDProyecto = Me.IDProyecto
                                                        .EdicionInicial = m_EdicionInicial
                                                        .EdicionFinal = m_EDicionFinal
                                                        .riesgo = m_ObjRiesgo.CodigoRiesgo
                                                        .NombreCampo = CStr(m_NombreCampo)
                                                        .Descripcion = "PCA" & m_NombreCampo & ":" & m_valorFinal & " " & m_NombreCampo & ":" & m_valorInicial
                                                        .Registrar p_Error
                                                        If p_Error <> "" Then
                                                            Err.Raise 1000
                                                        End If
                                                    End With
                                                    Set m_ObjCambio = Nothing
                                                Next
                                            End If
                                        End If
                                        
                                        Set m_ObjPCA = Nothing
                                        Set m_ObjPCAAnterior = Nothing
                                    Next
                                    For Each m_IDPCA In m_ObjColPCAAnterior.keys
                                        Set m_ObjPCAAnterior = m_ObjColPCAAnterior(m_IDPCA)
                                        Set m_ObjPCA = m_ObjPCAAnterior.PCAccionEdicionSiguiente
                                        p_Error = m_ObjPCAAnterior.Error
                                        If p_Error <> "" Then
                                            Err.Raise 1000
                                        End If
                                        If m_ObjPCA Is Nothing Then
                                            Set m_ObjCambio = New Cambio
                                            With m_ObjCambio
                                                .IDProyecto = Me.IDProyecto
                                                .EdicionInicial = m_EdicionInicial
                                                .EdicionFinal = m_EDicionFinal
                                                .riesgo = m_ObjRiesgo.CodigoRiesgo
                                                .Descripcion = "PCA" & " " & "Baja de Acción: " & m_ObjPCAAnterior.CodAccion & " del plan de contingencia: " & m_ObjPCAAnterior.Contingencia.CodContingencia
                                                .Registrar p_Error
                                                If p_Error <> "" Then
                                                    Err.Raise 1000
                                                End If
                                            End With
                                            Set m_ObjCambio = Nothing
                                            
                                        End If
                                        Set m_ObjPCAAnterior = Nothing
                                        Set m_ObjPCA = Nothing
                                    Next
                                Else
                                    'm_ObjColPCAAnterior Is Nothing
                                    For Each m_IDPCA In m_ObjColPCA.keys
                                        Set m_ObjPCA = m_ObjColPCA(m_IDPCA)
                                        Set m_ObjCambio = New Cambio
                                        With m_ObjCambio
                                            .IDProyecto = Me.IDProyecto
                                            .EdicionInicial = m_EdicionInicial
                                            .EdicionFinal = m_EDicionFinal
                                            .riesgo = m_ObjRiesgo.CodigoRiesgo
                                            .Descripcion = "PCA" & " " & "Alta de Acción: " & m_ObjPCA.CodAccion & " del plan de contingencia: " & m_ObjPCA.Contingencia.CodContingencia
                                            .Registrar p_Error
                                            If p_Error <> "" Then
                                                Err.Raise 1000
                                            End If
                                        End With
                                        Set m_ObjCambio = Nothing
                                    Next
                                End If
                            Else
                                If Not m_ObjColPCAAnterior Is Nothing Then
                                    For Each m_IDPCA In m_ObjColPCAAnterior.keys
                                        Set m_ObjPCAAnterior = m_ObjColPCAAnterior(m_IDPCA)
                                        Set m_ObjCambio = New Cambio
                                        With m_ObjCambio
                                            .IDProyecto = Me.IDProyecto
                                            .EdicionInicial = m_EdicionInicial
                                            .EdicionFinal = m_EDicionFinal
                                            .riesgo = m_ObjRiesgo.CodigoRiesgo
                                            .Descripcion = "PCA" & " " & "Baja de Acción: " & m_ObjPCAAnterior.CodAccion & " del plan de contingencia: " & m_ObjPCAAnterior.Contingencia.CodContingencia
                                            .Registrar p_Error
                                            If p_Error <> "" Then
                                                Err.Raise 1000
                                            End If
                                        End With
                                        Set m_ObjCambio = Nothing
                                        Set m_ObjPCAAnterior = Nothing
                                    Next
                                End If
                            End If
                            
                        End If
                        Set m_ObjPC = Nothing
                        Set m_ObjPCAnterior = Nothing
                    Next
                    For Each m_IDPC In m_ObjColPCAnterior.keys
                        Set m_ObjPCAnterior = m_ObjColPCAnterior(m_IDPC)
                        Set m_ObjPC = m_ObjPCAnterior.PCEdicionSiguiente
                        p_Error = m_ObjPCAnterior.Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        If m_ObjPC Is Nothing Then
                            Set m_ObjCambio = New Cambio
                            With m_ObjCambio
                                .IDProyecto = Me.IDProyecto
                                .EdicionInicial = m_EdicionInicial
                                .EdicionFinal = m_EDicionFinal
                                .riesgo = m_ObjRiesgo.CodigoRiesgo
                                .Descripcion = "PC" & " " & "Baja de plan de contingencia: " & m_ObjPCAnterior.CodContingencia
                                .Registrar p_Error
                                If p_Error <> "" Then
                                    Err.Raise 1000
                                End If
                            End With
                            Set m_ObjCambio = Nothing
                        End If
                    Next
                Else
                    'm_ObjColPCAnterior-->Nulo
                    For Each m_IDPC In m_ObjColPC.keys
                        Set m_ObjPC = m_ObjColPC(m_IDPC)
                        Set m_ObjCambio = New Cambio
                        With m_ObjCambio
                            .IDProyecto = Me.IDProyecto
                            .EdicionInicial = m_EdicionInicial
                            .EdicionFinal = m_EDicionFinal
                            .riesgo = m_ObjRiesgo.CodigoRiesgo
                            .Descripcion = "PC" & " " & "Alta de plan de contingencia: " & m_ObjPC.CodContingencia
                            .Registrar p_Error
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                        End With
                        Set m_ObjCambio = Nothing
                    Next
                End If
            Else
                If Not m_ObjColPCAnterior Is Nothing Then
                    For Each m_IDPC In m_ObjColPCAnterior.keys
                        Set m_ObjPCAnterior = m_ObjColPCAnterior(m_IDPC)
                        Set m_ObjCambio = New Cambio
                        With m_ObjCambio
                            .IDProyecto = Me.IDProyecto
                            .EdicionInicial = m_EdicionInicial
                            .EdicionFinal = m_EDicionFinal
                            .riesgo = m_ObjRiesgo.CodigoRiesgo
                            .Descripcion = "PC " & " " & "Baja de plan de contingencia: " & m_ObjPCAnterior.CodContingencia
                            .Registrar p_Error
                            If p_Error <> "" Then
                                Err.Raise 1000
                            End If
                        End With
                        Set m_ObjCambio = Nothing
                    Next
                End If
            End If
            
            
            
            Set m_ObjRiesgo = Nothing
            Set m_ObjRiesgoAnterior = Nothing
            Set m_objCambios = Nothing
            Set m_ObjColPM = Nothing
            Set m_ObjColPMAnterior = Nothing
            Set m_ObjPM = Nothing
            Set m_ObjPMAnterior = Nothing
            Set m_ObjColPMA = Nothing
            Set m_ObjColPMAAnterior = Nothing
            Set m_ObjPMA = Nothing
            Set m_ObjPMAAnterior = Nothing
            
            Set m_ObjColPC = Nothing
            Set m_ObjColPCAnterior = Nothing
            Set m_ObjPC = Nothing
            Set m_ObjPCAnterior = Nothing
            Set m_ObjColPCA = Nothing
            Set m_ObjColPCAAnterior = Nothing
            Set m_ObjPCA = Nothing
            Set m_ObjPCAAnterior = Nothing
siguienteRiesgo:
        Next
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarCambiosEnEdicion ha devuelto el error: " & vbNewLine & Err.Description
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
        .Add "IDEdicion"
        .Add "IDProyecto"
        .Add "FechaEdicion"
        .Add "Edicion"
        .Add "Elaborado"
        .Add "Revisado"
        .Add "Aprobado"
        .Add "FechaPublicacion"
        .Add "EntregadoAClienteORAC"
        .Add "Comentarios"
        .Add "PermitidoImprimirExcel"
        .Add "IDDocumentoAGEDO"
        .Add "NombreArchivoInforme"
        .Add "FechaMaxProximaPublicacion"
        .Add "FechaPreparadaParaPublicar"
        .Add "UsuarioProponePublicar"
        .Add "PropuestaRechazadaPorCalidadFecha"
        .Add "PropuestaRechazadaPorCalidadMotivo"
        .Add "UsuarioCalidadRechazaPropuesta"
        .Add "NotasCalidadParaPublicar"
        .Add "FechaUltimoCambio"
        .Add "UsuarioUltimoCambio"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
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
        .Add "Edicion"
        .Add "Elaborado"
        .Add "Revisado"
        .Add "Aprobado"
        .Add "EntregadoAClienteORAC"
        .Add "Comentarios"
        .Add "NombreArchivoInforme"
    End With
    Set ColCamposParaCambios = m_objColCamposParaCambios
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.ColCamposParaCambios ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property
Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDEdicion" Then
        getPropiedad = Me.IDEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        getPropiedad = Me.IDProyecto
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEdicion" Then
        getPropiedad = Me.FechaEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "Edicion" Then
        getPropiedad = Me.Edicion
        Exit Function
    End If
    If m_NombrePropiedad = "Elaborado" Then
        getPropiedad = Me.Elaborado
        Exit Function
    End If
    If m_NombrePropiedad = "Revisado" Then
        getPropiedad = Me.Revisado
        Exit Function
    End If
    If m_NombrePropiedad = "Aprobado" Then
        getPropiedad = Me.Aprobado
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPublicacion" Then
        getPropiedad = Me.FechaPublicacion
        Exit Function
    End If
    If m_NombrePropiedad = "EntregadoAClienteORAC" Then
        getPropiedad = Me.EntregadoAClienteORAC
        Exit Function
    End If
    If m_NombrePropiedad = "Comentarios" Then
        getPropiedad = Me.Comentarios
        Exit Function
    End If
    If m_NombrePropiedad = "PermitidoImprimirExcel" Then
        getPropiedad = Me.PermitidoImprimirExcel
        Exit Function
    End If
    If m_NombrePropiedad = "IDDocumentoAGEDO" Then
        getPropiedad = Me.IDDocumentoAGEDO
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivoInforme" Then
        getPropiedad = Me.NombreArchivoInforme
        Exit Function
    End If
    If m_NombrePropiedad = "FechaMaxProximaPublicacion" Then
        getPropiedad = Me.FechaMaxProximaPublicacion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPreparadaParaPublicar" Then
        getPropiedad = Me.FechaPreparadaParaPublicar
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioProponePublicar" Then
        getPropiedad = Me.UsuarioProponePublicar
        Exit Function
    End If
    If m_NombrePropiedad = "PropuestaRechazadaPorCalidadFecha" Then
        getPropiedad = Me.PropuestaRechazadaPorCalidadFecha
        Exit Function
    End If
    If m_NombrePropiedad = "PropuestaRechazadaPorCalidadMotivo" Then
        getPropiedad = Me.PropuestaRechazadaPorCalidadMotivo
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioCalidadRechazaPropuesta" Then
        getPropiedad = Me.UsuarioCalidadRechazaPropuesta
        Exit Function
    End If
    If m_NombrePropiedad = "NotasCalidadParaPublicar" Then
        getPropiedad = Me.NotasCalidadParaPublicar
        Exit Function
    End If
    If m_NombrePropiedad = "FechaUltimoCambio" Then
        getPropiedad = Me.FechaUltimoCambio
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioUltimoCambio" Then
        getPropiedad = Me.UsuarioUltimoCambio
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Edicion.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Edicion.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDEdicion" Then
        Me.IDEdicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        Me.IDProyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEdicion" Then
        Me.FechaEdicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Edicion" Then
        Me.Edicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Elaborado" Then
        Me.Elaborado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Revisado" Then
        Me.Revisado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Aprobado" Then
        Me.Aprobado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPublicacion" Then
        Me.FechaPublicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EntregadoAClienteORAC" Then
        Me.EntregadoAClienteORAC = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Comentarios" Then
        Me.Comentarios = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PermitidoImprimirExcel" Then
        Me.PermitidoImprimirExcel = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDDocumentoAGEDO" Then
        Me.IDDocumentoAGEDO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivoInforme" Then
        Me.NombreArchivoInforme = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaMaxProximaPublicacion" Then
        Me.FechaMaxProximaPublicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPreparadaParaPublicar" Then
        Me.FechaPreparadaParaPublicar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioProponePublicar" Then
        Me.UsuarioProponePublicar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PropuestaRechazadaPorCalidadFecha" Then
        Me.PropuestaRechazadaPorCalidadFecha = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PropuestaRechazadaPorCalidadMotivo" Then
        Me.PropuestaRechazadaPorCalidadMotivo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioCalidadRechazaPropuesta" Then
        Me.UsuarioCalidadRechazaPropuesta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NotasCalidadParaPublicar" Then
        Me.NotasCalidadParaPublicar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaUltimoCambio" Then
        Me.FechaUltimoCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioUltimoCambio" Then
        Me.UsuarioUltimoCambio = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Edicion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Edicion.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function




