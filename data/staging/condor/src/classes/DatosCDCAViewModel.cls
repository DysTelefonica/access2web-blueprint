VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosCDCAViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: DatosCDCAViewModel.cls (Normalizado Fase 2)
' RESPONSABILIDAD: Transporte de datos entre Capa de Datos y UI.
' ==========================================================================

' --- Entidades Principales ---
Public Solicitud As Solicitud
Public Expediente As Expediente
Public datos As DatosCDCA

' --- Datos de Estado y Flujo ---
Public EstadoNombre As String
Public TiempoCiclo As String

' --- Permisos de Edición por Bloque (Calculados por WorkflowServicio) ---
Public PuedeGuardar_Generales As Boolean
Public PuedeGuardar_Propuesta As Boolean
Public PuedeGuardar_Impacto As Boolean
Public PuedeGuardar_AprobacionSuministrador As Boolean
Public PuedeGuardar_DictamenRAC As Boolean
Public PuedeGuardar_DecisionFinal As Boolean

' --- Datos de Revisión de Calidad ---
Public vm_revisionCalidadEstado As String
Public vm_revisionCalidadComentarios As String

' ==========================================================================
' PROPIEDADES APLANADAS (Binding con Subformularios)
' ==========================================================================

' --- BLOQUE 1: GENERALES ---
Public vm_numContrato As String
Public vm_refSuministrador As String
Public vm_SuministradorNombreDir As String
Public vm_refDesviacionesPrevias As String
Public vm_requiereModificacionContrato As Variant ' Tri-estado (Sí/No/Null)

' --- BLOQUE 2: PROPUESTA TÉCNICA ---
Public vm_identificacionMaterial As String
Public vm_numPlanoEspecificacion As String
Public vm_cantidadPeriodo As String
Public vm_numSerieLote As String
Public vm_causaNC As String
Public vm_descripcionImpactoNC As String

' --- BLOQUE 3: IMPACTO ---
' Checkboxes (Booleanos)
Public vm_afectaPrestaciones As Boolean
Public vm_afectaSeguridad As Boolean
Public vm_afectaFiabilidad As Boolean
Public vm_afectaVidaUtil As Boolean
Public vm_afectaMedioambiente As Boolean
Public vm_afectaMantenibilidad As Boolean
Public vm_afectaIntercambiabilidad As Boolean
Public vm_afectaApariencia As Boolean
Public vm_afectaOtros As Boolean
' Selectores y Textos
Public vm_impactoCoste As String
Public vm_clasificacionNC As String
Public vm_esSuministradorAD As Variant ' Tri-estado
Public vm_identificacionAutoridadDiseno As String
Public vm_efectoFechaEntrega As String

' --- BLOQUE 4: APROBACIÓN SUMINISTRADOR ---
' Firmas Obligatorias
Public vm_firmaAprobacionRespIngenieriaNombre As String
Public vm_firmaAprobacionRespCalidadNombre As String
' Firmas Opcionales (Toggleables)
Public vm_firmaAprobacionRespProduccionNombre As String
Public vm_firmaAprobacionRespDisenioNombre As String
Public vm_firmaAprobacionRepresentanteSumNombre As String
' Estados de los Toggles (Interfaz)
Public vm_aplicaRespProduccion As Boolean
Public vm_aplicaRespDisenio As Boolean
Public vm_aplicaRepresentanteSum As Boolean

' --- BLOQUE 5: DICTAMEN RAC ---
Public vm_racCodigo As String
Public vm_racDecision As String
Public vm_observacionesRAC As String
Public vm_racNombre As String

' --- BLOQUE 6: DECISIÓN FINAL ---
Public vm_decisionFinal As String
Public vm_obsDecisionFinal As String
Public vm_NombreFirmanteFinal As String

' --- API uniforme de entidades CONDOR ---
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

