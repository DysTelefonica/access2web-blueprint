VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosPCViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ==========================================================================
' MÓDULO: DatosPCViewModel.cls
' ACCIÓN: Estandarizar nombres de variables RAC para coincidir con CDCA y BD.
' ==========================================================================
Option Compare Database
Option Explicit

' --- Núcleo del VM ---
Public idSolicitud As Long
Public datos As DatosPC
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

' --- Generales ---
Public vm_refContratoInspeccionOficial As String
Public vm_refSuministrador As String
Public vm_denominacionContrato As String
Public vm_SuministradorNombreDir As String
Public vm_objetoContrato As String

' --- Propuesta ---
Public vm_numPlanoEspecificacion As String
Public vm_descripcionMaterialAfectado As String
Public vm_descripcionPropuestaCambio As String

' --- Impacto ---
Public vm_motivoCorregirDeficiencias As Boolean
Public vm_motivoMejorarCapacidad As Boolean
Public vm_motivoAumentarNacionalizacion As Boolean
Public vm_motivoMejorarSeguridad As Boolean
Public vm_motivoMejorarFiabilidad As Boolean
Public vm_motivoMejorarCosteEficacia As Boolean
Public vm_motivoOtros As Boolean
Public vm_motivoOtrosDetalle As String

Public vm_incidenciaCoste As String
Public vm_incidenciaPlazo As String
Public vm_incidenciaSeguridad As Boolean
Public vm_incidenciaFiabilidad As Boolean
Public vm_incidenciaMantenibilidad As Boolean
Public vm_incidenciaIntercambiabilidad As Boolean
Public vm_incidenciaVidaUtilAlmacen As Boolean
Public vm_incidenciaFuncionamientoFuncion As Boolean
Public vm_impactoClasificacion As String
Public vm_CambioAfectaAMaterial As String

' --- Aprobación Suministrador ---
Public vm_firmaOficinaTecnicaNombre As String
Public vm_firmaRepSuministradorNombre As String

' --- Dictamen RAC (NOMBRE Normalizados) ---

Public vm_racDecision As String
Public vm_observacionesRAC As String   ' Antes: vm_dictamenRacObservaciones
Public vm_racNombre As String          ' Antes: vm_dictamenRACNombre
Public vm_racCodigo As String
' --- Decisión Final ---
Public vm_obsAprobacionAutoridadDiseno As String
Public vm_NombreAutoridadDiseno As String
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

