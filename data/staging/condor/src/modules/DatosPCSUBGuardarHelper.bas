Attribute VB_Name = "DatosPCSUBGuardarHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: DatosPCSUBGuardarHelper
' RESPONSABILIDAD: Plan testeable de guardado desde frmDatosPCSUB.
' ==========================================================================

Private Function NewGuardarResultado(ByVal p_Ok As Boolean) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "ok", p_Ok
    result.Add "bloque", 0
    result.Add "persisted", False
    result.Add "message", ""
    result.Add "navigationTarget", ""
    result.Add "workflowTarget", 0
    result.Add "requiresUserPrompt", False
    result.Add "errorText", ""
    Set NewGuardarResultado = result
End Function

Private Function BuildPCSUBGeneralesFromVm(ByRef p_ViewModel As DatosPCSUBViewModel) As DatosPCSUB
    Dim pcsub As New DatosPCSUB

    If p_ViewModel Is Nothing Then Err.Raise 91, "DatosPCSUBGuardarHelper", "ViewModel no inicializado."
    If p_ViewModel.Solicitud Is Nothing Then Err.Raise 91, "DatosPCSUBGuardarHelper", "Solicitud no inicializada."
    If Not p_ViewModel.datos Is Nothing Then pcsub.idDatosPCSUB = p_ViewModel.datos.idDatosPCSUB

    pcsub.idSolicitud = p_ViewModel.Solicitud.idSolicitud
    pcsub.refContratoInspeccionOficial = Nz(p_ViewModel.vm_refContratoInspeccionOficial, "")
    pcsub.refSubSuministrador = Nz(p_ViewModel.vm_refSubSuministrador, "")
    pcsub.SubsuministradorNombreDir = Nz(p_ViewModel.vm_SubsuministradorNombreDir, "")
    pcsub.denominacionContrato = Nz(p_ViewModel.vm_denominacionContrato, "")
    pcsub.objetoContrato = Nz(p_ViewModel.vm_objetoContrato, "")

    Set BuildPCSUBGeneralesFromVm = pcsub
End Function

Public Function DatosPCSUBGuardarHelper_GuardarDesdeSubformPlan( _
    ByVal p_NombreSubform As String, _
    ByRef p_ViewModel As DatosPCSUBViewModel, _
    Optional ByRef p_Db As DAO.Database = Nothing) As Object

    Dim result As Object
    Dim bloque As enumBloqueFormulario
    Dim pcsub As DatosPCSUB
    Dim svc As New DatosPCSUBServicio
    Dim transicionAutomatica As Boolean
    Dim solicitud As Solicitud

    On Error GoTo EH
    Set result = NewGuardarResultado(True)

    bloque = FormulariosPadreAuxiliares.MapearBloquePorNombreSub(p_NombreSubform, TipoForm_PCSUB)
    result("bloque") = CLng(bloque)
    If bloque <> Bloque_Generales Then Err.Raise 513, "DatosPCSUBGuardarHelper", "Este slice solo soporta Datos Generales PCSUB."

    Set solicitud = p_ViewModel.Solicitud
    Set pcsub = BuildPCSUBGeneralesFromVm(p_ViewModel)

    transicionAutomatica = svc.GuardarDatosGenerales(pcsub, p_Db)
    result("persisted") = True
    result("message") = "Datos Generales guardados."

    If bloque = Bloque_Generales And (solicitud.idEstadoInterno = estadoRegistro Or transicionAutomatica) Then
        result("requiresUserPrompt") = True
        If transicionAutomatica Then solicitud.idEstadoInterno = estadoRegistro
    End If

    Set DatosPCSUBGuardarHelper_GuardarDesdeSubformPlan = result
    Exit Function

EH:
    Set result = NewGuardarResultado(False)
    result("errorText") = Err.Description
    Set DatosPCSUBGuardarHelper_GuardarDesdeSubformPlan = result
End Function

Public Function DatosPCSUBGuardarHelper_AsignarTecnicoPostGuardado( _
    ByRef p_ViewModel As DatosPCSUBViewModel, _
    Optional ByRef p_Db As DAO.Database = Nothing) As Object

    Dim result As Object
    Dim wfServ As New WorkflowServicio
    Dim solicitud As Solicitud

    On Error GoTo EH
    Set result = NewGuardarResultado(True)

    If p_ViewModel Is Nothing Then Err.Raise 91, "DatosPCSUBGuardarHelper", "ViewModel no inicializado."
    If p_ViewModel.Solicitud Is Nothing Then Err.Raise 91, "DatosPCSUBGuardarHelper", "Solicitud no inicializada."

    Set solicitud = p_ViewModel.Solicitud
    If solicitud.idEstadoInterno <> estadoRegistro Then Err.Raise 513, "DatosPCSUBGuardarHelper", "La solicitud no está en Registro para asignación técnica."

    Call wfServ.EjecutarTransicion(solicitud, estadoDesarrolloTecnico, m_ObjUsuarioActivo, p_Db)
    solicitud.idEstadoInterno = estadoDesarrolloTecnico
    result("workflowTarget") = estadoDesarrolloTecnico
    result("message") = "Solicitud asignada a Desarrollo Técnico."

    Set DatosPCSUBGuardarHelper_AsignarTecnicoPostGuardado = result
    Exit Function

EH:
    Set result = NewGuardarResultado(False)
    result("errorText") = Err.Description
    Set DatosPCSUBGuardarHelper_AsignarTecnicoPostGuardado = result
End Function
