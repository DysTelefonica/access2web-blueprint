VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosPCSUBViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: DatosPCSUBViewModel.cls
' VIEWMODEL: Solicitud PC-Subcontratista
' FUENTE DE VERDAD: tbDatosPCSUB en condor_datos.accdb (ERD 2026-05-01)
' ==========================================================================

' --- Núcleo del VM ---
Public idSolicitud As Long
Public datos As DatosPCSUB
Public Solicitud As Solicitud
Public Expediente As Expediente

Public EstadoNombre As String
Public EstadoDescripcion As String
Public TiempoCiclo As String
Public TipoSolicitudNombre As String
Public CodigoNCAsociada As String

Public vm_revisionCalidadEstado As String
Public vm_revisionCalidadComentarios As String

' --- Flags de permiso ---
Public PuedeGuardar_Generales As Boolean
Public PuedeGuardar_Propuesta As Boolean
Public PuedeGuardar_Impacto As Boolean
Public PuedeGuardar_AprobacionSuministrador As Boolean
Public PuedeGuardar_DictamenRAC As Boolean
Public PuedeGuardar_DecisionFinal As Boolean

' --- Sección Generales ---
Public vm_refContratoInspeccionOficial As String
Public vm_denominacionContrato As String
Public vm_SubsuministradorNombreDir As String
Public vm_refSubSuministrador As String
Public vm_objetoContrato As String

' --- Propuesta ---
Public vm_numPlanoEspecificacion As String
Public vm_descripcionMaterialAfectado As String
Public vm_descripcionPropuestaCambio As String
Public vm_descripcionPropuestaCambioCont As String

' --- Impacto — Motivos ---
Public vm_motivoCorregirDeficiencias As Boolean
Public vm_motivoMejorarCapacidad As Boolean
Public vm_motivoAumentarNacionalizacion As Boolean
Public vm_motivoMejorarSeguridad As Boolean
Public vm_motivoMejorarFiabilidad As Boolean
Public vm_motivoMejorarCosteEficacia As Boolean
Public vm_motivoOtros As Boolean
Public vm_motivoOtrosDetalle As String

' --- Impacto — Incidencias ---
Public vm_incidenciaCoste As String
Public vm_incidenciaPlazo As String
Public vm_incidenciaSeguridad As Boolean
Public vm_incidenciaFiabilidad As Boolean
Public vm_incidenciaMantenibilidad As Boolean
Public vm_incidenciaIntercambiabilidad As Boolean
Public vm_incidenciaVidaUtilAlmacen As Boolean
Public vm_incidenciaFuncionamientoFuncion As Boolean

' --- Impacto — Clasificación ---
Public vm_impactoClasificacion As String
Public vm_CambioAfectaAMaterial As String

' --- Aprobación Suministrador Sub ---
Public vm_firmaOficinaTecnicaSubSuministradorNombre As String
Public vm_firmaRepSubSuministradorNombre As String

' --- Dictamen RAC ---
Public vm_racCodigo As String
Public vm_observacionesRAC As String
Public vm_racNombre As String
Public vm_racDecision As String
Public vm_racRechazoMotivos As String

' --- Dictamen RAC — Delegador ---
Public vm_observacionesRACDelegador As String
Public vm_racNombreDelegador As String

' --- Autoridad Diseño ---
Public vm_obsAprobacionAutoridadDiseno As String
Public vm_NombreAutoridadDiseno As String

' --- Decisión Final ---
Public vm_decisionFinal As String
Public vm_obsDecisionFinal As String
Public vm_NombreFirmanteFinal As String

' --- API uniforme ---
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
