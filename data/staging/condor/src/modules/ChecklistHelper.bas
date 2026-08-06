Attribute VB_Name = "ChecklistHelper"

Option Compare Database
Option Explicit

Private Function GetItemsModificacion(ByVal tipoSolicitud As String, Optional ByVal revisionCalidadEstado As String = "") As String
    Dim items As String
    items = ""
    
    items = items & "Validar la información indicada por el JP; completar los apartados Aprobación Suministrador y RAC."
    
    GetItemsModificacion = items
End Function

Private Function GetItemsDesarrolloTecnico(ByVal tipoSolicitud As String) As String
    Dim items As String
    items = ""
    items = items & "Completar en Desarrollo Técnico el formulario: apartados General, Detalle y Descripción, y Motivos (responsable: JP)."
    GetItemsDesarrolloTecnico = items
End Function

Private Function GetItemsRevision(ByVal tipoSolicitud As String) As String
    Dim items As String
    items = ""
    items = items & "Aprobación del RAC y formalización."
    GetItemsRevision = items
End Function

Private Function GetItemsFormalizacion(ByVal tipoSolicitud As String) As String
    Dim items As String
    items = ""
    items = items & "Cierre del expediente. La solicitud ha sido formalizada."
    GetItemsFormalizacion = items
End Function

Public Function GetItemsChecklistPorEstado( _
    ByVal idEstadoInterno As Long, _
    ByVal tipoSolicitud As String, _
    Optional ByVal revisionCalidadEstado As String = "", _
    Optional ByVal racValidacionEstado As String = "" _
) As String
    Dim items As String
    items = ""
    
    If revisionCalidadEstado = "RECHAZADO" And idEstadoInterno = estadoDesarrolloTecnico Then
        items = "Modificar el formulario con las indicaciones de Calidad."
    ElseIf revisionCalidadEstado = "RECHAZADO" And idEstadoInterno = estadoModificacion Then
        items = GetItemsModificacion(tipoSolicitud, revisionCalidadEstado)
    ElseIf idEstadoInterno = estadoRegistro Or idEstadoInterno = estadoDesarrolloTecnico Then
        items = GetItemsDesarrolloTecnico(tipoSolicitud)
    ElseIf idEstadoInterno = estadoModificacion Then
        items = GetItemsModificacion(tipoSolicitud, revisionCalidadEstado)
    ElseIf idEstadoInterno = estadoValidacion Then
        If racValidacionEstado = "RECHAZADO" Or racValidacionEstado = "COMENTARIOS" Then
            items = "Revisar el motivo del rechazo del RAC y tomar las acciones correspondientes."
        Else
            items = "El formulario ha sido aprobado. Próximo paso: envío del borrador al RAC."
        End If
    ElseIf idEstadoInterno = estadoRevision Then
        items = GetItemsRevision(tipoSolicitud)
    ElseIf idEstadoInterno = estadoFormalizacion Or idEstadoInterno = estadoAprobada Then
        items = GetItemsFormalizacion(tipoSolicitud)
    Else
        items = ""
    End If
    
    GetItemsChecklistPorEstado = items
End Function


