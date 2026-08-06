VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosCDCASUBViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public Solicitud As Solicitud
Public Expediente As Expediente
Public datos As DatosCDCASUB

Public EstadoNombre As String
Public TiempoCiclo As String

Public PuedeGuardar_Generales As Boolean
Public PuedeGuardar_Propuesta As Boolean
Public PuedeGuardar_Impacto As Boolean
Public PuedeGuardar_AprobacionSuministrador As Boolean
Public PuedeGuardar_DictamenRAC As Boolean
Public PuedeGuardar_DecisionFinal As Boolean

Public vm_revisionCalidadEstado As String
Public vm_revisionCalidadComentarios As String

' --- Aplanamiento ---
' Generales
Public vm_refSuministrador As String
Public vm_refSubSuministrador As String
Public vm_suministradorPrincipalNombreDir As String
Public vm_subSuministradorNombreDir As String
Public vm_refDesviacionesPrevias As String
Public vm_requiereModificacionContrato As Variant

' Propuesta
Public vm_identificacionMaterial As String
Public vm_numPlanoEspecificacion As String
Public vm_cantidadPeriodo As String
Public vm_numSerieLote As String
Public vm_causaNC As String
Public vm_descripcionImpactoNC As String

' Impacto
Public vm_afectaPrestaciones As Boolean
Public vm_afectaSeguridad As Boolean
Public vm_afectaFiabilidad As Boolean
Public vm_afectaVidaUtil As Boolean
Public vm_afectaMedioambiente As Boolean
Public vm_afectaMantenibilidad As Boolean
Public vm_afectaIntercambiabilidad As Boolean
Public vm_afectaApariencia As Boolean
Public vm_afectaOtros As Boolean
Public vm_impactoCoste As String
Public vm_clasificacionNC As String
Public vm_esSubSuministradorAD As Variant
Public vm_identificacionAutoridadDiseno As String
Public vm_efectoFechaEntrega As String

' Aprobacion
Public vm_firmaAprobacionRespIngenieriaNombre As String
Public vm_firmaAprobacionRespCalidadNombre As String
Public vm_firmaAprobacionRespProduccionNombre As String
Public vm_firmaAprobacionRespDisenioNombre As String
Public vm_firmaAprobacionRepresentanteSumNombre As String
' Toggles
Public vm_aplicaRespProduccion As Boolean
Public vm_aplicaRespDisenio As Boolean
Public vm_aplicaRepresentanteSum As Boolean

' RAC
Public vm_racCodigo As String
Public vm_racDecision As String
Public vm_observacionesRAC As String
Public vm_racNombre As String
' Extra CDCASUB
Public vm_observacionesRACDelegador As String
Public vm_racNombreDelegador As String

' Final
Public vm_decisionFinal As String
Public vm_obsDecisionFinal As String
Public vm_NombreFirmanteFinal As String

Public Function getPropiedad(ByVal nombre As String) As Variant
    On Error GoTo Fin
    getPropiedad = CallByName(Me, nombre, VbGet)
    Exit Function
Fin:
End Function

Public Sub SetPropiedad(ByVal nombre As String, ByVal valor As Variant)
    On Error Resume Next
    CallByName Me, nombre, VbLet, valor
End Sub

