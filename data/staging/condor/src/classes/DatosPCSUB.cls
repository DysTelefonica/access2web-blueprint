VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosPCSUB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: DatosPCSUB.cls
' ENTIDAD: Solicitud PC-Subcontratista (gemelo de PC + bloque subcontratista)
' FUENTE DE VERDAD: tbDatosPCSUB en condor_datos.accdb (ERD 2026-05-01)
' ==========================================================================

' --- PK y FK ---
Public idDatosPCSUB As Long
Public idSolicitud As Long

' --- Sección Generales ---
Public refContratoInspeccionOficial As String
Public refSubSuministrador As String
Public denominacionContrato As String
Public SubsuministradorNombreDir As String    ' Nota: S mayúscula (case-sensitive en VBA)
Public objetoContrato As String

' --- Sección Propuesta ---
Public descripcionMaterialAfectado As String
Public numPlanoEspecificacion As String
Public descripcionPropuestaCambio As String
Public descripcionPropuestaCambioCont As String

' --- Sección Impacto — Motivos ---
Public motivoCorregirDeficiencias As Boolean
Public motivoMejorarCapacidad As Boolean
Public motivoAumentarNacionalizacion As Boolean
Public motivoMejorarSeguridad As Boolean
Public motivoMejorarFiabilidad As Boolean
Public motivoMejorarCosteEficacia As Boolean
Public motivoOtros As Boolean
Public motivoOtrosDetalle As String

' --- Sección Impacto — Incidencias ---
Public incidenciaCoste As String
Public incidenciaPlazo As String
Public incidenciaSeguridad As Boolean
Public incidenciaFiabilidad As Boolean
Public incidenciaMantenibilidad As Boolean
Public incidenciaIntercambiabilidad As Boolean
Public incidenciaVidaUtilAlmacen As Boolean
Public incidenciaFuncionamientoFuncion As Boolean

' --- Sección Impacto — Clasificación ---
Public impactoClasificacion As String
Public CambioAfectaAMaterial As String

' --- Sección Aprobación Subcontratista ---
Public firmaOficinaTecnicaSubSuministradorNombre As String
Public firmaRepSubSuministradorNombre As String

' --- Sección Dictamen RAC (mismos campos que PC — no hay "racDelegado*") ---
Public racCodigo As String
Public observacionesRAC As String
Public racNombre As String
Public racDecision As String
Public racRechazoMotivos As String

' --- Sección Dictamen RAC — Delegador ---
Public observacionesRACDelegador As String
Public racNombreDelegador As String

' --- Sección Autoridad Diseño ---
Public obsAprobacionAutoridadDiseno As String
Public NombreAutoridadDiseno As String

' --- Sección Decisión Final ---
Public decisionFinal As String
Public obsDecisionFinal As String
Public NombreFirmanteFinal As String

' --- Colección de campos para reflexión ---
Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            ' PK/FK
            .Add "idDatosPCSUB"
            .Add "idSolicitud"

            ' Sección Generales
            .Add "refContratoInspeccionOficial"
            .Add "refSubSuministrador"
            .Add "denominacionContrato"
            .Add "SubsuministradorNombreDir"
            .Add "objetoContrato"

            ' Sección Propuesta
            .Add "descripcionMaterialAfectado"
            .Add "numPlanoEspecificacion"
            .Add "descripcionPropuestaCambio"
            .Add "descripcionPropuestaCambioCont"

            ' Sección Impacto — Motivos
            .Add "motivoCorregirDeficiencias"
            .Add "motivoMejorarCapacidad"
            .Add "motivoAumentarNacionalizacion"
            .Add "motivoMejorarSeguridad"
            .Add "motivoMejorarFiabilidad"
            .Add "motivoMejorarCosteEficacia"
            .Add "motivoOtros"
            .Add "motivoOtrosDetalle"

            ' Sección Impacto — Incidencias
            .Add "incidenciaCoste"
            .Add "incidenciaPlazo"
            .Add "incidenciaSeguridad"
            .Add "incidenciaFiabilidad"
            .Add "incidenciaMantenibilidad"
            .Add "incidenciaIntercambiabilidad"
            .Add "incidenciaVidaUtilAlmacen"
            .Add "incidenciaFuncionamientoFuncion"

            ' Sección Impacto — Clasificación
            .Add "impactoClasificacion"
            .Add "CambioAfectaAMaterial"

            ' Sección Aprobación
            .Add "firmaOficinaTecnicaSubSuministradorNombre"
            .Add "firmaRepSubSuministradorNombre"

            ' Sección Dictamen RAC
            .Add "racCodigo"
            .Add "observacionesRAC"
            .Add "racNombre"
            .Add "racDecision"
            .Add "racRechazoMotivos"

            ' Sección Dictamen RAC — Delegador
            .Add "observacionesRACDelegador"
            .Add "racNombreDelegador"

            ' Sección Autoridad Diseño
            .Add "obsAprobacionAutoridadDiseno"
            .Add "NombreAutoridadDiseno"

            ' Sección Decisión Final
            .Add "decisionFinal"
            .Add "obsDecisionFinal"
            .Add "NombreFirmanteFinal"
        End With
    End If
    Set ColCampos = m_objColCampos
End Property

Public Function getPropiedad(ByVal m_NombrePropiedad As String) As Variant
    getPropiedad = CallByName(Me, m_NombrePropiedad, VbGet)
End Function

Public Sub SetPropiedad(ByVal m_NombrePropiedad As String, ByVal m_ValorPropiedad As Variant)
    CallByName Me, m_NombrePropiedad, VbLet, m_ValorPropiedad
End Sub
